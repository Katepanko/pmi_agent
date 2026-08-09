type RuntimeEnv = { DB?: D1Database };

async function database() {
  const { env } = await import("cloudflare:workers");
  const db = (env as RuntimeEnv).DB;
  if (!db) throw new Error("Persistent storage is unavailable: the DB binding is not configured.");
  return db;
}

let schemaReady: Promise<void> | null = null;

export async function ensureCoreSchema() {
  if (schemaReady) return schemaReady;
  const db = await database();
  schemaReady = db.batch([
    db.prepare(`CREATE TABLE IF NOT EXISTS projects (
      id TEXT PRIMARY KEY NOT NULL, user_id TEXT NOT NULL, name TEXT NOT NULL,
      description TEXT NOT NULL DEFAULT '', companies TEXT NOT NULL DEFAULT '',
      deal_rationale TEXT NOT NULL DEFAULT '', integration_type TEXT NOT NULL DEFAULT '',
      day_one_date TEXT, objectives TEXT NOT NULL DEFAULT '', synergy_targets TEXT NOT NULL DEFAULT '',
      governance TEXT NOT NULL DEFAULT '', terminology TEXT NOT NULL DEFAULT '',
      reporting_expectations TEXT NOT NULL DEFAULT '', instructions TEXT NOT NULL DEFAULT '',
      icon TEXT NOT NULL DEFAULT 'layers', archived_at TEXT, created_at TEXT NOT NULL, updated_at TEXT NOT NULL
    )`),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_projects_user_archived ON projects(user_id, archived_at)"),
    db.prepare(`CREATE TABLE IF NOT EXISTS chats (
      id TEXT PRIMARY KEY NOT NULL, user_id TEXT NOT NULL, project_id TEXT,
      title TEXT NOT NULL, model_key TEXT NOT NULL DEFAULT 'openai-gpt56',
      audience TEXT NOT NULL DEFAULT 'Steering Committee', archived_at TEXT,
      created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
      FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
    )`),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_chats_user_project ON chats(user_id, project_id)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_chats_user_updated ON chats(user_id, updated_at)"),
    db.prepare(`CREATE TABLE IF NOT EXISTS messages (
      id TEXT PRIMARY KEY NOT NULL, chat_id TEXT NOT NULL, role TEXT NOT NULL,
      content TEXT NOT NULL, classification TEXT NOT NULL DEFAULT 'conversation',
      model_key TEXT, source_coverage_json TEXT, stopped_at TEXT, created_at TEXT NOT NULL,
      FOREIGN KEY(chat_id) REFERENCES chats(id) ON DELETE CASCADE
    )`),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_messages_chat_created ON messages(chat_id, created_at)"),
    db.prepare(`CREATE TABLE IF NOT EXISTS sources (
      id TEXT PRIMARY KEY NOT NULL, user_id TEXT NOT NULL, project_id TEXT, chat_id TEXT NOT NULL,
      message_id TEXT, file_name TEXT NOT NULL, file_type TEXT NOT NULL, media_type TEXT NOT NULL,
      size_bytes INTEGER NOT NULL, object_key TEXT NOT NULL, checksum TEXT,
      extraction_status TEXT NOT NULL DEFAULT 'pending', extraction_warnings_json TEXT NOT NULL DEFAULT '[]',
      metadata_json TEXT NOT NULL DEFAULT '{}', created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
      FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
      FOREIGN KEY(chat_id) REFERENCES chats(id) ON DELETE CASCADE,
      FOREIGN KEY(message_id) REFERENCES messages(id) ON DELETE SET NULL
    )`),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_sources_chat_status ON sources(chat_id, extraction_status)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_sources_project ON sources(project_id)"),
  ]).then(() => undefined).catch((error) => {
    schemaReady = null;
    throw error;
  });
  return schemaReady;
}

export function authenticatedUserId(request: Request) {
  return request.headers.get("oai-authenticated-user-id") ?? "local-preview";
}

export type WorkspaceSnapshot = {
  projects: Array<{ id: string; name: string; monogram: string; context: string; chats: string[]; expanded: boolean }>;
  chats: Array<{
    id: string;
    title: string;
    audience: string;
    modelKey?: string;
    projectId: string | null;
    messages: Array<{ id: string; role: "user" | "assistant"; content: string; createdAt: string; variant?: "demo-report" | "error" }>;
    sources: Array<{ id: string; name: string; type: string; size: number; status: "extracted" | "partial" | "pending" | "failed"; excerpt?: string; warnings?: string[] }>;
  }>;
};

export async function loadWorkspace(userId: string): Promise<WorkspaceSnapshot> {
  await ensureCoreSchema();
  const db = await database();
  const [projectRows, chatRows, messageRows, sourceRows] = await db.batch([
    db.prepare("SELECT id, name, description, icon FROM projects WHERE user_id = ? AND archived_at IS NULL ORDER BY updated_at DESC").bind(userId),
    db.prepare("SELECT id, project_id, title, model_key, audience FROM chats WHERE user_id = ? AND archived_at IS NULL ORDER BY updated_at DESC").bind(userId),
    db.prepare("SELECT m.id, m.chat_id, m.role, m.content, m.classification, m.created_at FROM messages m INNER JOIN chats c ON c.id = m.chat_id WHERE c.user_id = ? ORDER BY m.created_at ASC").bind(userId),
    db.prepare("SELECT id, project_id, chat_id, file_name, file_type, size_bytes, extraction_status, extraction_warnings_json, metadata_json FROM sources WHERE user_id = ? ORDER BY created_at ASC").bind(userId),
  ]);

  const rawProjects = projectRows.results as Array<Record<string, unknown>>;
  const rawChats = chatRows.results as Array<Record<string, unknown>>;
  const rawMessages = messageRows.results as Array<Record<string, unknown>>;
  const rawSources = sourceRows.results as Array<Record<string, unknown>>;

  const chats = rawChats.map((chat) => ({
    id: String(chat.id),
    title: String(chat.title),
    audience: String(chat.audience),
    modelKey: String(chat.model_key),
    projectId: chat.project_id ? String(chat.project_id) : null,
    messages: rawMessages.filter((message) => message.chat_id === chat.id).map((message) => ({
      id: String(message.id),
      role: message.role as "user" | "assistant",
      content: String(message.content),
      createdAt: new Date(String(message.created_at)).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }),
      variant: message.classification === "error" ? "error" as const : undefined,
    })),
    sources: rawSources.filter((source) => source.chat_id === chat.id).map((source) => {
      const metadata = JSON.parse(String(source.metadata_json || "{}")) as { excerpt?: string };
      return {
        id: String(source.id),
        name: String(source.file_name),
        type: String(source.file_type),
        size: Number(source.size_bytes),
        status: source.extraction_status as "extracted" | "partial" | "pending" | "failed",
        excerpt: metadata.excerpt,
        warnings: JSON.parse(String(source.extraction_warnings_json || "[]")) as string[],
      };
    }),
  }));

  return {
    projects: rawProjects.map((project) => ({
      id: String(project.id),
      name: String(project.name),
      monogram: String(project.icon || "PM").slice(0, 2).toUpperCase(),
      context: String(project.description || ""),
      chats: chats.filter((chat) => chat.projectId === project.id).map((chat) => chat.id),
      expanded: true,
    })),
    chats,
  };
}

export async function syncWorkspace(userId: string, snapshot: WorkspaceSnapshot) {
  await ensureCoreSchema();
  const db = await database();
  const now = new Date().toISOString();

  for (const project of snapshot.projects) {
    const existing = await db.prepare("SELECT user_id FROM projects WHERE id = ?").bind(project.id).first<{ user_id: string }>();
    if (existing && existing.user_id !== userId) throw new Error("Project boundary violation.");
    await db.prepare(`INSERT INTO projects (id, user_id, name, description, icon, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(id) DO UPDATE SET name = excluded.name, description = excluded.description, icon = excluded.icon, updated_at = excluded.updated_at
    `).bind(project.id, userId, project.name, project.context, project.monogram, now, now).run();
  }

  for (const chat of snapshot.chats) {
    const existing = await db.prepare("SELECT user_id FROM chats WHERE id = ?").bind(chat.id).first<{ user_id: string }>();
    if (existing && existing.user_id !== userId) throw new Error("Chat boundary violation.");
    if (chat.projectId) {
      const owner = await db.prepare("SELECT user_id FROM projects WHERE id = ?").bind(chat.projectId).first<{ user_id: string }>();
      if (!owner || owner.user_id !== userId) throw new Error("Chat project boundary violation.");
    }
    await db.prepare(`INSERT INTO chats (id, user_id, project_id, title, model_key, audience, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(id) DO UPDATE SET title = excluded.title, model_key = excluded.model_key, audience = excluded.audience, updated_at = excluded.updated_at
    `).bind(chat.id, userId, chat.projectId, chat.title, chat.modelKey ?? "openai-gpt56", chat.audience, now, now).run();

    for (const message of chat.messages) {
      await db.prepare(`INSERT INTO messages (id, chat_id, role, content, classification, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET content = excluded.content, classification = excluded.classification
      `).bind(message.id, chat.id, message.role, message.content, message.variant === "error" ? "error" : "conversation", now).run();
    }

    for (const source of chat.sources) {
      const sourceOwner = await db.prepare("SELECT user_id FROM sources WHERE id = ?").bind(source.id).first<{ user_id: string }>();
      if (sourceOwner && sourceOwner.user_id !== userId) throw new Error("Source boundary violation.");
      await db.prepare(`INSERT INTO sources (id, user_id, project_id, chat_id, file_name, file_type, media_type, size_bytes, object_key, extraction_status, extraction_warnings_json, metadata_json, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET extraction_status = excluded.extraction_status, extraction_warnings_json = excluded.extraction_warnings_json, metadata_json = excluded.metadata_json, updated_at = excluded.updated_at
      `).bind(source.id, userId, chat.projectId, chat.id, source.name, source.type, "application/octet-stream", source.size, `staged/${userId}/${source.id}`, source.status, JSON.stringify(source.warnings ?? []), JSON.stringify({ excerpt: source.excerpt }), now, now).run();
    }
  }
  await db.prepare("PRAGMA optimize").run();
}
