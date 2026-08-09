"use client";

import { DragEvent, FormEvent, KeyboardEvent, useEffect, useRef, useState } from "react";
import { createReportArtifact, type ReportFormat } from "./lib/report-export";

type ModelOption = {
  key: string;
  displayName: string;
  provider: string;
  available: boolean;
  unavailableReason?: string;
};

type Attachment = {
  id: string;
  name: string;
  type: string;
  size: number;
  status: "extracted" | "partial" | "pending" | "failed";
  excerpt?: string;
  warnings?: string[];
};

type Message = {
  id: string;
  role: "user" | "assistant";
  content: string;
  createdAt: string;
  attachments?: Attachment[];
  variant?: "demo-report" | "error";
  artifact?: { name: string; format: ReportFormat; url: string };
};

type Chat = {
  id: string;
  title: string;
  audience: string;
  modelKey?: string;
  projectId: string | null;
  messages: Message[];
  sources: Attachment[];
};

type Project = {
  id: string;
  name: string;
  monogram: string;
  context: string;
  chats: string[];
  expanded: boolean;
};

const demoFiles: Attachment[] = [
  { id: "src-masterplan", name: "Integration_Masterplan.xlsx", type: "xlsx", size: 284_000, status: "extracted", excerpt: "Structured milestone and workstream records with cell-level provenance." },
  { id: "src-risks", name: "Risk_Register.xlsx", type: "xlsx", size: 126_000, status: "extracted", excerpt: "Risk status, impact, mitigation, owners, and due dates." },
  { id: "src-synergies", name: "Synergy_Tracker.xlsx", type: "xlsx", size: 194_000, status: "extracted", excerpt: "Targets, actuals, forecasts, timing, and initiative owners." },
  { id: "src-weekly", name: "Weekly_Update.docx", type: "docx", size: 88_000, status: "extracted", excerpt: "Workstream updates, decisions, actions, and blockers." },
];

const demoMessages: Message[] = [
  {
    id: "demo-user",
    role: "user",
    createdAt: "09:41",
    content: "Create a concise SteerCo report covering integration trajectory, critical risks, synergy delivery, and decisions required.",
    attachments: demoFiles,
  },
  {
    id: "demo-assistant",
    role: "assistant",
    createdAt: "09:42",
    content: "Illustrative report preview",
    variant: "demo-report",
  },
];

const initialChats: Chat[] = [
  {
    id: "executive-report",
    title: "Executive report",
    audience: "Steering Committee",
    projectId: "orion",
    messages: demoMessages,
    sources: demoFiles,
  },
  { id: "finance-review", title: "Finance review", audience: "CFO", projectId: "orion", messages: [], sources: [] },
  { id: "weekly-imo", title: "Weekly IMO", audience: "IMO / PMO", projectId: "orion", messages: [], sources: [] },
  { id: "day-one", title: "Day 1 readiness", audience: "Steering Committee", projectId: "atlas", messages: [], sources: [] },
  { id: "synergy-review", title: "Synergy review", audience: "Board", projectId: "atlas", messages: [], sources: [] },
  { id: "standalone", title: "Quick analysis", audience: "Infer from request", projectId: null, messages: [], sources: [] },
];

const initialProjects: Project[] = [
  {
    id: "orion",
    name: "Orion / Nova",
    monogram: "ON",
    context: "Orion acquired Nova. Day 1: 1 October 2026. Partial integration focused on Finance, IT, HR, and commercial value capture. Weekly IMO and monthly SteerCo reporting.",
    chats: ["executive-report", "finance-review", "weekly-imo"],
    expanded: true,
  },
  {
    id: "atlas",
    name: "Atlas Health",
    monogram: "AH",
    context: "Confidential healthcare integration. Standalone project boundary; no Orion / Nova knowledge may be used.",
    chats: ["day-one", "synergy-review"],
    expanded: false,
  },
];

const audienceOptions = ["Board", "CEO", "Steering Committee", "IMO / PMO", "CFO", "Workstream team", "Infer from request"];

function uid(prefix: string) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
}

function fileSize(size: number) {
  if (size < 1024) return `${size} B`;
  if (size < 1_048_576) return `${Math.round(size / 1024)} KB`;
  return `${(size / 1_048_576).toFixed(1)} MB`;
}

function fileGlyph(type: string) {
  if (["xlsx", "xls", "csv"].includes(type)) return "XL";
  if (type === "pptx") return "PP";
  if (type === "pdf") return "PD";
  if (type === "docx") return "WD";
  if (["png", "jpg", "jpeg"].includes(type)) return "IM";
  return "FL";
}

function requestedFormat(value: string): ReportFormat | null {
  const text = value.toLowerCase();
  if (!/\b(generate|create|export|download|make)\b/.test(text)) return null;
  if (/powerpoint|pptx|slide deck/.test(text)) return "pptx";
  if (/excel|xlsx|spreadsheet/.test(text)) return "xlsx";
  if (/word|docx|document/.test(text)) return "docx";
  if (/pdf/.test(text)) return "pdf";
  if (/html|dashboard/.test(text)) return "html";
  return null;
}

export function PMIWorkspace({ initialModels }: { initialModels: ModelOption[] }) {
  const [projects, setProjects] = useState(initialProjects);
  const [chats, setChats] = useState(initialChats);
  const [activeChatId, setActiveChatId] = useState("executive-report");
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [contextOpen, setContextOpen] = useState(false);
  const [sourceOpen, setSourceOpen] = useState(false);
  const [composer, setComposer] = useState("");
  const [queuedFiles, setQueuedFiles] = useState<Attachment[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [workspaceDirty, setWorkspaceDirty] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const [selectedModel, setSelectedModel] = useState(initialModels.find((model) => model.key === "openai-gpt56")?.key ?? initialModels[0]?.key ?? "openai-gpt56");
  const abortRef = useRef<AbortController | null>(null);
  const textAreaRef = useRef<HTMLTextAreaElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const conversationRef = useRef<HTMLDivElement | null>(null);

  const activeChat = chats.find((chat) => chat.id === activeChatId) ?? chats[0];
  const activeProject = projects.find((project) => project.id === activeChat.projectId) ?? null;
  const selectedModelOption = initialModels.find((model) => model.key === selectedModel);

  useEffect(() => {
    let cancelled = false;
    void fetch("/api/workspace")
      .then(async (response) => response.ok ? response.json() : Promise.reject(new Error("Workspace unavailable")))
      .then((snapshot: { projects?: Project[]; chats?: Chat[] }) => {
        if (cancelled || !snapshot.chats?.length) return;
        setProjects(snapshot.projects ?? []);
        setChats(snapshot.chats);
        setActiveChatId(snapshot.chats[0].id);
        if (snapshot.chats[0].modelKey) setSelectedModel(snapshot.chats[0].modelKey);
      })
      .catch(() => undefined);
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!workspaceDirty) return;
    const timer = window.setTimeout(() => {
      void fetch("/api/workspace", {
        method: "PUT",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ projects, chats }),
      }).then((response) => { if (response.ok) setWorkspaceDirty(false); }).catch(() => undefined);
    }, 700);
    return () => window.clearTimeout(timer);
  }, [workspaceDirty, projects, chats]);

  const allSourceCount = activeChat.sources.length;
  const coverageComplete = allSourceCount > 0 && activeChat.sources.every((source) => source.status === "extracted");

  const scrollToBottom = () => {
    requestAnimationFrame(() => conversationRef.current?.scrollTo({ top: conversationRef.current.scrollHeight, behavior: "smooth" }));
  };

  const updateActiveChat = (updater: (chat: Chat) => Chat) => {
    setWorkspaceDirty(true);
    setChats((current) => current.map((chat) => (chat.id === activeChatId ? updater(chat) : chat)));
  };

  const downloadReport = (format: ReportFormat, content?: string, addMessage = true) => {
    const draft = content ?? [...activeChat.messages].reverse().find((message) => message.role === "assistant" && message.variant !== "error" && message.content.trim())?.content;
    if (!draft) return false;
    const { blob, fileName } = createReportArtifact(format, {
      title: activeChat.title,
      audience: activeChat.audience,
      content: draft,
      sources: activeChat.sources.map((source) => ({ name: source.name, status: source.status })),
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    if (addMessage) {
      updateActiveChat((chat) => ({ ...chat, messages: [...chat.messages, {
        id: uid("message"), role: "assistant", createdAt: "Now",
        content: `Your ${format.toUpperCase()} file is ready.`,
        artifact: { name: fileName, format, url },
      }] }));
      scrollToBottom();
    }
    setExportOpen(false);
    return true;
  };

  const addFiles = async (files: File[]) => {
    const supported = ["xlsx", "xls", "csv", "pptx", "docx", "pdf", "html", "htm", "png", "jpg", "jpeg"];
    const additions = await Promise.all(
      files.map(async (file): Promise<Attachment> => {
        const type = file.name.split(".").pop()?.toLowerCase() || "file";
        const textReadable = ["csv", "html", "htm"].includes(type);
        const excerpt = textReadable ? (await file.text()).slice(0, 4_000) : undefined;
        return {
          id: uid("file"),
          name: file.name,
          type,
          size: file.size,
          status: supported.includes(type) ? "pending" : "failed",
          excerpt,
        };
      }),
    );
    setQueuedFiles((current) => [...current, ...additions]);

    const accepted = files.filter((file) => supported.includes(file.name.split(".").pop()?.toLowerCase() || ""));
    if (!accepted.length) return;
    await Promise.all(accepted.map(async (file, index) => {
      const attachment = additions[files.indexOf(file)] ?? additions[index];
      try {
        const form = new FormData();
        form.append("files", file);
        form.append("fileId", attachment.id);
        const response = await fetch("/api/extract", { method: "POST", body: form });
        const payload = (await response.json()) as {
          documents?: Array<{ fileId: string; status: Attachment["status"]; rawText: string; extractionWarnings: string[] }>;
          error?: string;
        };
        const document = payload.documents?.[0];
        if (!response.ok || !document) throw new Error(payload.error ?? "Extraction service unavailable.");
        setQueuedFiles((current) => current.map((queued) => queued.id === attachment.id
          ? { ...queued, status: document.status, excerpt: document.rawText.slice(0, 4_000), warnings: document.extractionWarnings }
          : queued));
      } catch (error) {
        const warning = error instanceof Error ? error.message : "Extraction service unavailable.";
        setQueuedFiles((current) => current.map((queued) => queued.id === attachment.id
          ? { ...queued, status: "pending", warnings: [`${warning} Retry this source without re-uploading the other files.`] }
          : queued));
      }
    }));
  };

  const handleDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragging(false);
    void addFiles(Array.from(event.dataTransfer.files));
  };

  const resizeComposer = () => {
    const area = textAreaRef.current;
    if (!area) return;
    area.style.height = "0px";
    area.style.height = `${Math.min(Math.max(area.scrollHeight, 56), 176)}px`;
  };

  const stopGeneration = () => {
    abortRef.current?.abort();
    abortRef.current = null;
    setIsGenerating(false);
    updateActiveChat((chat) => {
      const messages = [...chat.messages];
      const last = messages.at(-1);
      if (last?.role === "assistant" && !last.content.trim()) {
        messages[messages.length - 1] = { ...last, content: "Generation stopped." };
      }
      return { ...chat, messages };
    });
  };

  const submitMessage = async (event?: FormEvent) => {
    event?.preventDefault();
    const messageText = composer.trim();
    if ((!messageText && queuedFiles.length === 0) || isGenerating) return;

    const userMessage: Message = {
      id: uid("message"),
      role: "user",
      content: messageText || "Please review the attached files.",
      createdAt: new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }),
      attachments: queuedFiles,
    };
    const assistantId = uid("message");
    const combinedSources = [...activeChat.sources, ...queuedFiles];
    const history = activeChat.messages.map(({ role, content }) => ({ role, content }));
    const format = requestedFormat(userMessage.content);
    const latestDraft = [...activeChat.messages].reverse().find((message) => message.role === "assistant" && message.variant !== "error" && message.content.trim())?.content;

    if (format && latestDraft && queuedFiles.length === 0) {
      updateActiveChat((chat) => ({ ...chat, messages: [...chat.messages, userMessage] }));
      setComposer("");
      if (textAreaRef.current) textAreaRef.current.style.height = "56px";
      window.setTimeout(() => downloadReport(format, latestDraft), 0);
      return;
    }

    updateActiveChat((chat) => ({
      ...chat,
      modelKey: selectedModel,
      messages: [...chat.messages, userMessage, { id: assistantId, role: "assistant", content: "", createdAt: "Now" }],
      sources: combinedSources,
      title: chat.messages.length === 0 ? messageText.slice(0, 42) || "Source review" : chat.title,
    }));
    setComposer("");
    setQueuedFiles([]);
    if (textAreaRef.current) textAreaRef.current.style.height = "56px";
    setIsGenerating(true);
    scrollToBottom();

    const controller = new AbortController();
    abortRef.current = controller;

    try {
      const response = await fetch("/api/chat", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          modelKey: selectedModel,
          message: userMessage.content,
          history,
          audience: activeChat.audience,
          projectContext: activeProject?.context,
          sources: combinedSources.map((source) => ({
            id: source.id,
            fileName: source.name,
            status: source.status,
            excerpt: source.excerpt,
            warnings: source.warnings ?? (source.status === "pending" ? ["Binary extraction is queued; coverage is not yet complete."] : []),
          })),
        }),
        signal: controller.signal,
      });

      if (!response.ok) {
        const payload = (await response.json().catch(() => ({}))) as { error?: string };
        throw new Error(payload.error ?? "The selected model could not be reached.");
      }
      if (!response.body) throw new Error("The model returned no response stream.");

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        const token = decoder.decode(value, { stream: true });
        updateActiveChat((chat) => ({
          ...chat,
          messages: chat.messages.map((message) =>
            message.id === assistantId ? { ...message, content: message.content + token } : message,
          ),
        }));
        scrollToBottom();
      }
    } catch (error) {
      if ((error as Error).name !== "AbortError") {
        const detail = error instanceof Error ? error.message : "Generation failed.";
        updateActiveChat((chat) => ({
          ...chat,
          messages: chat.messages.map((message) =>
            message.id === assistantId
              ? {
                  ...message,
                  variant: "error",
                  content: `Model connection required\n\n${detail}\n\nAdd the provider key and configured model ID to the server environment. Uploaded files and project context remain in this chat.`,
                }
              : message,
          ),
        }));
      }
    } finally {
      abortRef.current = null;
      setIsGenerating(false);
      scrollToBottom();
    }
  };

  const onComposerKeyDown = (event: KeyboardEvent<HTMLTextAreaElement>) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      void submitMessage();
    }
  };

  const createChat = (projectId: string | null) => {
    const id = uid("chat");
    const chat: Chat = { id, title: "New conversation", audience: "Infer from request", projectId, messages: [], sources: [] };
    setChats((current) => [chat, ...current]);
    setWorkspaceDirty(true);
    if (projectId) {
      setProjects((current) => current.map((project) => project.id === projectId ? { ...project, expanded: true, chats: [id, ...project.chats] } : project));
    }
    setActiveChatId(id);
  };

  const createProject = () => {
    const id = uid("project");
    const chatId = uid("chat");
    setProjects((current) => [{ id, name: "Untitled integration", monogram: "UI", context: "Add project context to help every project chat start with the same facts and terminology.", chats: [chatId], expanded: true }, ...current]);
    setChats((current) => [{ id: chatId, title: "Project kickoff", audience: "Infer from request", projectId: id, messages: [], sources: [] }, ...current]);
    setWorkspaceDirty(true);
    setActiveChatId(chatId);
    setContextOpen(true);
  };

  const renameChat = (chat: Chat) => {
    const title = window.prompt("Rename chat", chat.title)?.trim();
    if (!title || title === chat.title) return;
    setChats((current) => current.map((item) => item.id === chat.id ? { ...item, title } : item));
    setWorkspaceDirty(true);
  };

  const deleteChat = (chat: Chat) => {
    if (!window.confirm(`Delete “${chat.title}”? This removes its messages and sources.`)) return;
    const remaining = chats.filter((item) => item.id !== chat.id);
    if (remaining.length === 0) {
      const replacement: Chat = { id: uid("chat"), title: "New conversation", audience: "Infer from request", projectId: null, messages: [], sources: [] };
      setChats([replacement]);
      setActiveChatId(replacement.id);
    } else {
      setChats(remaining);
      if (activeChatId === chat.id) setActiveChatId(remaining[0].id);
    }
    setProjects((current) => current.map((project) => ({ ...project, chats: project.chats.filter((id) => id !== chat.id) })));
    setWorkspaceDirty(true);
  };

  const renameProject = (project: Project) => {
    const name = window.prompt("Rename project", project.name)?.trim();
    if (!name || name === project.name) return;
    const monogram = name.split(/\s+/).slice(0, 2).map((word) => word[0]).join("").toUpperCase();
    setProjects((current) => current.map((item) => item.id === project.id ? { ...item, name, monogram: monogram || "PM" } : item));
    setWorkspaceDirty(true);
  };

  const deleteProject = (project: Project) => {
    if (!window.confirm(`Delete “${project.name}” and all of its chats?`)) return;
    const removed = new Set(project.chats);
    const remaining = chats.filter((chat) => !removed.has(chat.id));
    if (remaining.length === 0) {
      const replacement: Chat = { id: uid("chat"), title: "New conversation", audience: "Infer from request", projectId: null, messages: [], sources: [] };
      setChats([replacement]);
      setActiveChatId(replacement.id);
    } else {
      setChats(remaining);
      if (removed.has(activeChatId)) setActiveChatId(remaining[0].id);
    }
    setProjects((current) => current.filter((item) => item.id !== project.id));
    setWorkspaceDirty(true);
  };

  const toggleProject = (projectId: string) => {
    setProjects((current) => current.map((project) => project.id === projectId ? { ...project, expanded: !project.expanded } : project));
  };

  const updateProjectContext = (context: string) => {
    if (!activeProject) return;
    setProjects((current) => current.map((project) => project.id === activeProject.id ? { ...project, context } : project));
    setWorkspaceDirty(true);
  };

  const updateAudience = (audience: string) => {
    updateActiveChat((chat) => ({ ...chat, audience }));
  };

  const currentModelReason = selectedModelOption?.unavailableReason;

  return (
    <div className={`app-shell ${sidebarCollapsed ? "sidebar-is-collapsed" : ""}`}>
      <aside className="sidebar" aria-label="Projects and chats">
        <div className="brand-row">
          <button className="brand" onClick={() => setSidebarCollapsed(false)} aria-label="PMI Agent home">
            <span className="brand-mark">P</span>
            {!sidebarCollapsed && <span className="brand-name">PMI AGENT</span>}
          </button>
          {!sidebarCollapsed && <button className="icon-button inverse" onClick={() => setSidebarCollapsed(true)} aria-label="Collapse sidebar">‹</button>}
        </div>

        <div className="primary-actions">
          <button className="new-chat-button" onClick={() => createChat(null)}>
            <span>＋</span><span className="sidebar-label">New chat</span><kbd>⌘ K</kbd>
          </button>
          <button className="new-project-button" onClick={createProject}>
            <span className="project-plus">◇</span><span className="sidebar-label">New project</span>
          </button>
        </div>

        {!sidebarCollapsed && (
          <div className="sidebar-scroll">
            <div className="sidebar-section-label">Projects</div>
            <div className="project-list">
              {projects.map((project) => (
                <div className="project-block" key={project.id}>
                  <div className="project-row">
                    <button className="project-toggle" onClick={() => toggleProject(project.id)} aria-expanded={project.expanded}>
                      <span className="project-monogram">{project.monogram}</span>
                      <span className="project-name">{project.name}</span>
                      <span className="chevron">{project.expanded ? "⌄" : "›"}</span>
                    </button>
                    <button className="project-add" onClick={() => createChat(project.id)} aria-label={`New chat in ${project.name}`}>＋</button>
                    <button className="sidebar-action" onClick={() => renameProject(project)} aria-label={`Rename ${project.name}`}>✎</button>
                    <button className="sidebar-action danger" onClick={() => deleteProject(project)} aria-label={`Delete ${project.name}`}>×</button>
                  </div>
                  {project.expanded && (
                    <div className="nested-chats">
                      {project.chats.map((chatId) => {
                        const chat = chats.find((item) => item.id === chatId);
                        if (!chat) return null;
                        return (
                          <div key={chat.id} className={`chat-row ${activeChatId === chat.id ? "active" : ""}`}>
                            <button className="chat-link" onClick={() => setActiveChatId(chat.id)}><span className="chat-rail" /><span>{chat.title}</span></button>
                            <button className="sidebar-action" onClick={() => renameChat(chat)} aria-label={`Rename ${chat.title}`}>✎</button>
                            <button className="sidebar-action danger" onClick={() => deleteChat(chat)} aria-label={`Delete ${chat.title}`}>×</button>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              ))}
            </div>

            <div className="sidebar-section-label standalone-label">Standalone chats</div>
            {chats.filter((chat) => chat.projectId === null).map((chat) => (
              <div key={chat.id} className={`chat-row standalone-chat ${activeChatId === chat.id ? "active" : ""}`}>
                <button className="chat-link" onClick={() => setActiveChatId(chat.id)}><span className="bubble-icon">○</span><span>{chat.title}</span></button>
                <button className="sidebar-action" onClick={() => renameChat(chat)} aria-label={`Rename ${chat.title}`}>✎</button>
                <button className="sidebar-action danger" onClick={() => deleteChat(chat)} aria-label={`Delete ${chat.title}`}>×</button>
              </div>
            ))}
          </div>
        )}

        <div className="sidebar-footer">
          <div className="profile-avatar">U</div>
          {!sidebarCollapsed && <div><strong>User</strong><span>Private workspace</span></div>}
          {!sidebarCollapsed && <button className="footer-more" aria-label="Workspace menu">•••</button>}
        </div>
      </aside>

      <main className="workspace">
        <header className="chat-header">
          <div className="header-title-group">
            {sidebarCollapsed && <button className="icon-button sidebar-open" onClick={() => setSidebarCollapsed(false)} aria-label="Open sidebar">›</button>}
            <div>
              <div className="eyebrow">{activeProject ? activeProject.name : "Standalone chat"}</div>
              <h1>{activeChat.title}</h1>
            </div>
          </div>

          <div className="header-controls">
            <label className="model-select-wrap">
              <span className={`provider-dot ${selectedModelOption?.provider ?? "openai"}`} />
              <select value={selectedModel} onChange={(event) => setSelectedModel(event.target.value)} aria-label="AI model">
                {initialModels.map((model) => <option key={model.key} value={model.key}>{model.displayName}{model.available ? "" : " · setup"}</option>)}
              </select>
              <span className="select-chevron">⌄</span>
            </label>
            <button className={`header-pill ${sourceOpen ? "active" : ""}`} onClick={() => setSourceOpen((open) => !open)}>
              <span className="stack-icon">▱</span> Sources <b>{allSourceCount}</b>
            </button>
            {activeProject && <button className={`header-pill context-button ${contextOpen ? "active" : ""}`} onClick={() => setContextOpen((open) => !open)}>Project context</button>}
            <div className="export-wrap">
              <button className={`header-pill ${exportOpen ? "active" : ""}`} onClick={() => setExportOpen((open) => !open)}>Export <span>⌄</span></button>
              {exportOpen && <div className="export-menu">
                {(["pptx", "pdf", "xlsx", "docx", "html"] as ReportFormat[]).map((format) => <button key={format} onClick={() => downloadReport(format)}>{format === "pptx" ? "PowerPoint" : format === "xlsx" ? "Excel workbook" : format === "docx" ? "Word document" : format === "html" ? "HTML dashboard" : "PDF"}<span>.{format}</span></button>)}
              </div>}
            </div>
            <button className="more-button" aria-label="Chat actions">•••</button>
          </div>
        </header>

        <div className="content-stage">
          <section className="conversation" ref={conversationRef} aria-live="polite">
            {activeChat.messages.length === 0 ? (
              <EmptyConversation project={activeProject} onPrompt={setComposer} />
            ) : (
              <div className="message-column">
                <div className="conversation-date"><span>Today</span></div>
                {activeChat.messages.map((message) => (
                  <MessageView key={message.id} message={message} generating={isGenerating && message.id === activeChat.messages.at(-1)?.id} />
                ))}
              </div>
            )}
          </section>

          <div
            className={`composer-zone ${isDragging ? "is-dragging" : ""}`}
            onDragEnter={(event) => { event.preventDefault(); setIsDragging(true); }}
            onDragOver={(event) => event.preventDefault()}
            onDragLeave={(event) => { if (event.currentTarget === event.target) setIsDragging(false); }}
            onDrop={handleDrop}
          >
            <form className="composer" onSubmit={submitMessage}>
              {queuedFiles.length > 0 && (
                <div className="queued-files">
                  {queuedFiles.map((file) => (
                    <div className="queued-file" key={file.id}>
                      <span className={`file-glyph ${file.type}`}>{fileGlyph(file.type)}</span>
                      <span><strong>{file.name}</strong><small>{fileSize(file.size)}</small></span>
                      <button type="button" onClick={() => setQueuedFiles((current) => current.filter((item) => item.id !== file.id))} aria-label={`Remove ${file.name}`}>×</button>
                    </div>
                  ))}
                </div>
              )}
              <textarea
                ref={textAreaRef}
                value={composer}
                onChange={(event) => { setComposer(event.target.value); resizeComposer(); }}
                onKeyDown={onComposerKeyDown}
                placeholder="Ask about your integration or request a management report…"
                aria-label="Message"
                rows={1}
              />
              <div className="composer-toolbar">
                <div className="composer-tools">
                  <button type="button" className="attach-button" onClick={() => fileInputRef.current?.click()} aria-label="Attach files">⌕<span>Attach</span></button>
                  <input ref={fileInputRef} type="file" multiple accept=".xlsx,.xls,.csv,.pptx,.docx,.pdf,.html,.htm,.png,.jpg,.jpeg" onChange={(event) => { void addFiles(Array.from(event.target.files ?? [])); event.target.value = ""; }} />
                  <span className="divider" />
                  <label className="audience-select">
                    <span>Audience</span>
                    <select value={activeChat.audience} onChange={(event) => updateAudience(event.target.value)} aria-label="Report audience">
                      {audienceOptions.map((audience) => <option key={audience}>{audience}</option>)}
                    </select>
                  </label>
                </div>
                {isGenerating ? (
                  <button type="button" className="send-button stop-button" onClick={stopGeneration} aria-label="Stop generation"><span /></button>
                ) : (
                  <button type="submit" className="send-button" disabled={!composer.trim() && queuedFiles.length === 0} aria-label="Send message">↑</button>
                )}
              </div>
            </form>
            <div className="composer-meta">
              <span className={`status-dot ${selectedModelOption?.available ? "connected" : ""}`} />
              {selectedModelOption?.available ? `${selectedModelOption.displayName} connected` : currentModelReason ?? "Provider setup required"}
              <span className="meta-separator">•</span>
              <span>Every applicable source is checked before synthesis</span>
            </div>
            {isDragging && <div className="drop-overlay"><span>＋</span>Drop files into this message</div>}
          </div>

          {(sourceOpen || contextOpen) && (
            <aside className="detail-drawer" aria-label={sourceOpen ? "Sources" : "Project context"}>
              <div className="drawer-header">
                <div><span className="eyebrow">Evidence workspace</span><h2>{sourceOpen ? "Sources" : "Project context"}</h2></div>
                <button className="drawer-close" onClick={() => { setSourceOpen(false); setContextOpen(false); }} aria-label="Close panel">×</button>
              </div>
              {sourceOpen ? (
                <SourceDrawer sources={activeChat.sources} coverageComplete={coverageComplete} />
              ) : activeProject ? (
                <ContextDrawer project={activeProject} onChange={updateProjectContext} />
              ) : null}
            </aside>
          )}
        </div>
      </main>
    </div>
  );
}

function EmptyConversation({ project, onPrompt }: { project: Project | null; onPrompt: (value: string) => void }) {
  const prompts = [
    "Create a Board-ready integration status report",
    "Identify conflicts across all uploaded sources",
    "Summarize red and amber risks and decisions required",
  ];
  return (
    <div className="empty-conversation">
      <div className="empty-mark">P</div>
      <span className="eyebrow">PMI reporting agent</span>
      <h2>Turn integration evidence into a management point of view.</h2>
      <p>{project ? `This chat inherits context from ${project.name}. ` : "This is a standalone chat. "}Attach trackers, updates, decks, minutes, or screenshots, then ask for the deliverable you need.</p>
      <div className="prompt-suggestions">
        {prompts.map((prompt) => <button key={prompt} onClick={() => onPrompt(prompt)}>{prompt}<span>↗</span></button>)}
      </div>
    </div>
  );
}

function MessageView({ message, generating }: { message: Message; generating: boolean }) {
  if (message.role === "user") {
    return (
      <article className="message user-message">
        {message.attachments && message.attachments.length > 0 && (
          <div className="message-files">
            {message.attachments.map((file) => (
              <div className="message-file" key={file.id}>
                <span className={`file-glyph ${file.type}`}>{fileGlyph(file.type)}</span>
                <span><strong>{file.name}</strong><small>{file.type.toUpperCase()} · {fileSize(file.size)}</small></span>
                <span className={`file-status ${file.status}`}>{file.status === "extracted" ? "Indexed" : file.status === "pending" ? "Queued" : file.status}</span>
              </div>
            ))}
          </div>
        )}
        <div className="user-bubble">{message.content}</div>
        <time>{message.createdAt}</time>
      </article>
    );
  }

  return (
    <article className={`message assistant-message ${message.variant === "error" ? "has-error" : ""}`}>
      <div className="assistant-avatar">P</div>
      <div className="assistant-body">
        <div className="assistant-meta"><strong>PMI Agent</strong><span>PMI consultant</span><time>{message.createdAt}</time></div>
        {message.variant === "demo-report" ? <DemoReport /> : (
          <div className="streamed-content">{message.content}{generating && <span className="typing-cursor" />}</div>
        )}
        {message.artifact && <a className="artifact-download" href={message.artifact.url} download={message.artifact.name}><span className={`file-glyph ${message.artifact.format}`}>{fileGlyph(message.artifact.format)}</span><span><strong>{message.artifact.name}</strong><small>Download {message.artifact.format.toUpperCase()}</small></span><b>↓</b></a>}
        {!generating && message.content && message.variant !== "error" && (
          <div className="message-actions"><button>Copy</button><button>Useful</button><button>Needs work</button><button>•••</button></div>
        )}
      </div>
    </article>
  );
}

function DemoReport() {
  return (
    <div className="report-preview">
      <div className="report-kicker"><span>Draft v1</span><span>Steering Committee</span><span>Illustrative demo</span></div>
      <h2>Integration momentum is broadly intact, but two cross-functional dependencies require SteerCo intervention</h2>
      <p className="lead">The current evidence indicates that delivery remains manageable if ownership and timing are resolved this week. Value capture is progressing, while the most material execution exposure is concentrated at the IT–Finance boundary.</p>

      <div className="insight-grid">
        <div className="insight-card">
          <span className="card-index">01</span><h3>Trajectory</h3>
          <p>Most planned workstreams are advancing, with current pressure concentrated in activities that depend on shared systems and cutover sequencing.</p>
        </div>
        <div className="insight-card">
          <span className="card-index">02</span><h3>Value delivery</h3>
          <p>Synergy initiatives remain active; the draft separates confirmed source values from forecast assumptions and flags timing gaps for validation.</p>
        </div>
        <div className="insight-card critical">
          <span className="card-index">03</span><h3>Management focus</h3>
          <p>Two decisions need named owners and dates before the next reporting cycle to protect the integrated plan.</p>
        </div>
      </div>

      <h3 className="section-heading">Decisions required</h3>
      <div className="decision-list">
        <div><span className="rag red">R</span><div><strong>Confirm the accountable executive for the cross-functional cutover dependency</strong><p>Management implication: unresolved accountability prevents a credible recovery path from being reflected in the masterplan.</p></div><span className="decision-timing">This week</span></div>
        <div><span className="rag amber">A</span><div><strong>Validate whether the latest synergy timing supersedes the baseline tracker</strong><p>A source inconsistency should be resolved before external value-delivery reporting.</p></div><span className="decision-timing">Before SteerCo</span></div>
      </div>

      <div className="ai-recommendation">
        <span className="ai-label">AI-generated recommendation — validation required</span>
        <p>Run a focused IT–Finance dependency review before the next SteerCo and publish one reconciled cutover baseline with a single accountable owner.</p>
        <div className="evidence-row"><strong>Evidence basis</strong><span>Dependency records</span><span>Weekly workstream update</span><span>Ownership gap</span></div>
      </div>

      <div className="coverage-strip">
        <div className="coverage-icon">✓</div>
        <div><strong>Source coverage complete</strong><span>4 of 4 applicable files considered · material conflicts and gaps disclosed</span></div>
        <button>Inspect evidence →</button>
      </div>
      <div className="draft-note"><span>Content preview first</span>The report remains editable in chat. Generate PowerPoint, Excel, Word, or PDF only after the draft is ready.</div>
    </div>
  );
}

function SourceDrawer({ sources, coverageComplete }: { sources: Attachment[]; coverageComplete: boolean }) {
  if (sources.length === 0) return <div className="drawer-empty"><span>▱</span><h3>No sources yet</h3><p>Attach files in any message. Every applicable source will be included in the coverage check.</p></div>;
  return (
    <div className="drawer-content">
      <div className={`coverage-summary ${coverageComplete ? "complete" : "incomplete"}`}>
        <span>{coverageComplete ? "✓" : "!"}</span>
        <div><strong>{coverageComplete ? "Coverage complete" : "Extraction pending"}</strong><small>{sources.filter((source) => source.status === "extracted").length} of {sources.length} sources ready</small></div>
      </div>
      <div className="source-list">
        {sources.map((source) => (
          <button className="source-row" key={source.id}>
            <span className={`file-glyph ${source.type}`}>{fileGlyph(source.type)}</span>
            <span className="source-details"><strong>{source.name}</strong><small>{source.excerpt ?? "Awaiting deterministic extraction"}</small></span>
            <span className={`source-state ${source.status}`}>{source.status}</span>
          </button>
        ))}
      </div>
      <div className="source-priority">
        <span className="eyebrow">Source authority</span>
        <h3>No priority rule recorded</h3>
        <p>Tell the agent, for example: “Use the masterplan as authoritative for milestone dates.” The rule will persist in this chat.</p>
      </div>
    </div>
  );
}

function ContextDrawer({ project, onChange }: { project: Project; onChange: (value: string) => void }) {
  const [value, setValue] = useState(project.context);
  return (
    <div className="drawer-content context-drawer">
      <div className="context-project-mark">{project.monogram}</div>
      <h3>{project.name}</h3>
      <p>Project context is inherited by every chat in this project and isolated from all other projects.</p>
      <label><span>Integration context</span><textarea value={value} onChange={(event) => setValue(event.target.value)} rows={10} /></label>
      <button className="save-context" onClick={() => onChange(value)}>Save context</button>
      <div className="boundary-note"><strong>Knowledge boundary</strong><span>Project ID: {project.id}</span><p>Standalone chats and other projects do not inherit this context.</p></div>
    </div>
  );
}
