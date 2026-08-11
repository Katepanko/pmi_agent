# How This Agent Was Generated

This document explains both how the PMI Reporting Agent application was created
and how an answer or report is generated when someone uses it.

## First: non-technical view

### Where it started

The agent started with a detailed written product brief in `prompt.txt`. The brief
described the intended user experience, the types of PMI information the agent
should understand, the reports it should produce, and the safeguards it should
follow.

The central idea was simple: a consultant uploads fragmented integration
material, asks a management question, and receives a decision-ready answer or
report based on all available evidence.

The repository was otherwise empty at the beginning. The first working version
was generated from that brief on top of a standard full-stack web application
starter. It was then improved in several iterations: first the chat and file
processing foundation, then more resilient uploads and workspace management,
then PowerPoint generation, and finally Word, Excel, PDF, and HTML reports.

### What was actually generated

This project did **not** train a new language model from scratch and did not
fine-tune a model on PMI data. Instead, it created a specialist application
around an existing OpenAI or Anthropic model.

The specialization comes from five layers:

1. **A PMI role and working method.** The model is instructed to behave like a
   senior Post-Merger Integration consultant and to write for Boards, Steering
   Committees, IMOs, CFOs, and workstream teams.
2. **Project and conversation context.** The application supplies the selected
   audience, project description, recent conversation, and previous report
   version with each request.
3. **Evidence from uploaded files.** The application extracts information from
   spreadsheets, presentations, documents, web exports, PDFs, and images. It
   keeps the source name and location attached to extracted content.
4. **Evidence controls.** Every applicable source is included in a coverage
   check. Conflicting values, incomplete extraction, and missing information
   must be disclosed instead of silently resolved or invented.
5. **Professional output builders.** When a user requests a report, the model
   plans the content and storyline. Deterministic software then turns that plan
   into an editable PowerPoint, Excel workbook, Word document, PDF, or HTML
   dashboard.

### How a user request becomes an answer

When a user sends a message, the agent follows this flow:

1. It combines the message with the project context, intended audience, recent
   chat history, and uploaded sources.
2. It checks whether every source has been extracted completely, partially, or
   not at all.
3. It compares evidence across sources and identifies disagreements that need
   to remain visible.
4. It asks the selected language model to synthesize the evidence using the PMI
   consulting instructions.
5. For a normal question, it streams the answer into the chat.
6. For a requested deliverable, it creates a structured report plan, validates
   it, renders the requested file, stores it, and returns a download link.
7. If the user requests a revision, the prior structured report is loaded so
   unaffected sections can be preserved in a new version.

### What makes it a PMI agent

The underlying language model provides general reasoning and writing ability.
The application makes that ability useful for PMI work by supplying domain
instructions, complete source context, management-reporting conventions,
conflict handling, provenance, and output formats. In other words, the “agent”
is the full workflow, not only the model call.

### Current boundaries

The application is intentionally transparent where coverage is incomplete.
Spreadsheet, presentation, Word, CSV, and HTML structure can be extracted
directly. PDF extraction is conservative, legacy `.xls` files require a future
compatibility adapter, and images or embedded slide visuals require multimodal
inspection before their meaning can be considered fully covered.

## Second: technical view

### 1. Generation input and implementation history

The primary generation specification is `prompt.txt`. It defines the product,
provider abstraction, security expectations, supported inputs, source-coverage
rules, PMI reasoning standards, conversation behavior, artifact requirements,
and phased implementation plan.

`docs/architecture-audit.md` records that the workspace initially contained only
that specification. The implementation history shows the following sequence:

- the initial conversational application, data model, extraction pipeline,
  provider adapters, and UI;
- resilient uploads, workspace behavior, and early exports;
- model-planned PowerPoint creation and persisted artifact versions;
- multi-format PPTX, XLSX, DOCX, PDF, and HTML artifact generation;
- storage-binding, validation, evidence-conflict, and branding improvements.

The current application uses React 19, TypeScript, a Next-compatible App Router,
vinext, and a Cloudflare Worker-compatible runtime.

### 2. Runtime composition of the agent

The runtime system prompt is assembled in `app/lib/pmi-prompt.ts`. Its main
inputs are:

- the PMI system instruction;
- audience and project context;
- user-defined source-authority rules;
- a complete source manifest;
- deterministic cross-source reconciliation;
- the current draft, when one exists;
- a final validation instruction covering evidence, gaps, calculations,
  conflicts, audience fit, and recommendation labels.

The separate Markdown files in `prompts/` document the intended prompt modules
for system behavior, PMI consulting, grounding, synthesis, validation, and
format-specific report generation. The active chat path currently consolidates
the core behavior in the TypeScript prompt builders.

This is prompt-based orchestration, not model training. The provider receives a
system prompt plus user/assistant messages and returns either a stream of text
or a complete structured-artifact plan.

### 3. Model-provider abstraction

`app/lib/models.ts` builds a model registry from server-side environment
variables. Display keys are kept separate from provider model IDs so the UI does
not depend on a provider's naming scheme.

`app/lib/llm/provider.ts` defines the common provider contract. The OpenAI and
Anthropic adapters implement that contract in `app/lib/llm/openai.ts` and
`app/lib/llm/anthropic.ts`. `app/lib/llm/index.ts` selects the adapter at
runtime. API keys remain in server-side environment variables and are not sent
to the browser.

### 4. File ingestion and provenance

The browser posts attachments to `POST /api/extract`. The endpoint:

- accepts up to 20 files per message and 25 MB per file;
- stages original binaries in the `FILES` object-storage binding when present;
- invokes the deterministic extraction pipeline;
- returns compact extracted content, warnings, coverage status, and evidence
  reconciliation.

`app/lib/ingestion/pipeline.ts` routes by extension and preserves useful source
locations:

- CSV: headers and row-level records;
- XLSX: sheets, rows, cell references, values, and formulas;
- DOCX: paragraphs plus basic transcript signals;
- PPTX: slide text and speaker notes;
- HTML: semantic blocks;
- PDF: a conservative low-confidence text-layer pass;
- PNG/JPEG: dimensions and a marker that visual interpretation is pending;
- XLS: partial status with a compatibility warning.

Each extracted document receives an ID, extraction status, metadata, structured
elements, raw text, and warnings. The coverage check is complete only when all
applicable documents have an `extracted` status.

### 5. Evidence reconciliation and grounding

`app/lib/evidence.ts` extracts comparable observations from source excerpts and
normalizes candidate metrics, values, dates, scopes, and authority signals. It
groups observations into facts and classifies disagreements. A value may be
selected only when supported by the reconciliation rules; otherwise the fact
remains an unresolved conflict.

`app/lib/conflict-guard.ts` enforces visibility of material conflicts in
structured artifacts. The deterministic reconciliation is inserted into the
model prompt and is described as mandatory, so the language model is not asked
to improvise conflict resolution.

This design separates two responsibilities:

- deterministic code establishes what sources exist, their extraction state,
  and whether evidence conflicts;
- the language model synthesizes meaning, implications, storyline, and clearly
  labelled recommendations.

### 6. Chat request path

The main orchestration endpoint is `POST /api/chat` in
`app/api/chat/route.ts`.

For every request it:

1. validates the message and selected configured model;
2. reconciles evidence from all supplied sources, recent user statements, and
   source-authority rules;
3. identifies whether the request is ordinary chat or an artifact request;
4. follows the text or artifact branch.

For text chat, `buildGroundedPrompt()` composes the system instruction and the
selected provider streams the response. The endpoint includes up to the most
recent 20 conversation messages.

For artifact generation, the provider first returns a constrained JSON content
model. The application parses and limits that structure, enforces conflict
visibility, renders it, validates the generated binary, uploads it to object
storage, and records its metadata in the database.

### 7. Artifact generation

`app/lib/artifact-intent.ts` detects requested output formats. Format planning
is implemented in `app/lib/artifact.ts`; PowerPoint has additional storyline and
layout logic in `app/lib/presentation.ts`.

The model owns semantic decisions such as:

- governing thought and storyline;
- section or slide selection;
- message-led headings;
- facts, implications, gaps, and recommendations;
- evidence references and management audience.

The renderers own deterministic file construction:

- `app/lib/presentation.ts` for PPTX;
- `app/lib/renderers/excel.ts` for XLSX;
- `app/lib/renderers/word.ts` for DOCX;
- `app/lib/renderers/pdf.ts` for PDF;
- `app/lib/renderers/html.ts` for HTML.

Generated artifacts are checked for a non-empty payload, correct extension, and
format-specific package structure. Additional checks verify expected branding
assets in the current branded renderers.

### 8. Persistence and deployment

Structured state is stored in Cloudflare D1/SQLite through the `DB` binding.
Original uploads and generated artifacts are stored in R2-compatible object
storage through the `FILES` binding.

The schema includes projects, chats, messages, sources, extracted segments,
source-priority rules, knowledge versions, drafts, artifact versions,
provenance, and model-usage records. Artifact records retain the structured
model, version number, and parent artifact ID, enabling conversational revision
without rebuilding every unaffected part from scratch.

The application is built for a Cloudflare Worker-compatible ESM deployment via
vinext. Standard deployment and binding configuration lives in `wrangler.jsonc`
and `vite.config.ts`; runtime binding access is isolated in the binding helpers.
The application does not depend on OpenAI hosting.

### 9. Validation

The repository includes build and Node test scripts for rendered HTML,
presentations, multi-format artifacts, and evidence conflicts. The standard
validation commands are:

```bash
npm test
npm run lint
npm run db:generate
```

These checks validate the software paths and generated file structures. They do
not replace human review of consulting judgments, which is why generated
recommendations are labelled as requiring validation.
