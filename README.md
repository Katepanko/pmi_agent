# PMI Agent — Decision-ready integration reporting

PMI Agent is a conversational Post-Merger Integration reporting workspace. It
combines project context, heterogeneous source material, source coverage, and
provider-independent language models to produce evidence-grounded management
report drafts.

## Current release

The current vertical slice includes:

- projects, project chats, and standalone chats;
- isolated project context and audience selection;
- multi-file first messages, drag-and-drop, attachment state, streaming, and
  cancellation;
- configurable OpenAI and Anthropic model registry and provider adapters;
- deterministic extraction for CSV, HTML, XLSX, DOCX, and PPTX text/structure;
- conservative partial extraction for PDFs and images so incomplete coverage is
  never silently presented as complete;
- row, sheet, slide, paragraph, and HTML-block provenance;
- source coverage inspection and extraction warnings;
- D1-backed project, chat, message, and source metadata persistence;
- version-ready knowledge, authority-rule, draft, draft-version, provenance, and
  model-usage schema;
- explicit PowerPoint intent and revision handling with a structured, LLM-owned storyline model;
- server-side consulting-style PPTX rendering, R2 storage, D1 artifact metadata, and per-message downloads;
- responsive black, green, and white consulting-chat interface.

See [docs/architecture-audit.md](docs/architecture-audit.md) for the initial
audit and staged gap map.

See [docs/how-this-agent-was-generated.md](docs/how-this-agent-was-generated.md)
for a non-technical and technical explanation of how the application was built
and how it generates answers and reports.

## Local setup

Requires Node.js 22.13 or later.

```bash
npm install
cp .env.example .env.local
npm run dev
```

Provider keys and actual model IDs stay server-side. Both are required for a
model to become available in the selector. The UI displays a clear setup state
instead of fabricating a response when configuration is incomplete.

## Validation

```bash
npm test
npm run lint
npm run db:generate
```

The deployment build targets Cloudflare Workers through vinext. D1 is declared
as `DB`; R2 is declared as `FILES`. The generated SQL migration is stored in
`drizzle/`.

## Honest limitations

This release establishes the conversational, ingestion, grounding, and
persistence foundation. The next production phases are:

- a layout-aware PDF parser and multimodal inspection for source images and
  slide visuals;
- a legacy `.xls` compatibility adapter;
- hierarchical retrieval over large corpora and deterministic conflict
  normalization;
- full knowledge-version and source-priority-rule services;
- editable XLSX/DOCX/PDF renderers and configured corporate-template ingestion;
- authentication policy, enterprise tenancy, malware scanning, retention,
  observability, and provider cost estimation.

Those gaps are surfaced in extraction status and documentation rather than
represented as complete.
