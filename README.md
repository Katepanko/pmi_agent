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
- optional `@filename` template selection that separates template content from deal evidence, preflights semantic layout/capacity, fills PPTX, XLSX, DOCX, PDF, and HTML outputs, accepts image/CSV references, and leaves ordinary no-template prompts unchanged;
- server-side consulting-style PPTX rendering, R2 storage, D1 artifact metadata, and per-message downloads;
- responsive black, green, and white consulting-chat interface.

See [docs/architecture-audit.md](docs/architecture-audit.md) for the initial
audit and staged gap map.

See [docs/how-this-agent-was-generated.md](docs/how-this-agent-was-generated.md)
for a non-technical and technical explanation of how the application was built
and how it generates answers and reports.

## Using a report template

Attach the template and the PMI evidence files in the same chat. Click the `@`
button on the template attachment (or type its exact `@filename`) and ask for
the new report, for example: `Using @Nike_adidas_PMI_Executive_Status.pptx,
create the same report for our PMI deal.` The marked file controls the report
structure and is excluded from deal evidence; every other attachment remains an
evidence source. Prompts without an `@filename` continue to use the standard
artifact workflow.

Template generation is structure-first: the server inspects the selected file
before asking the model for content, maps evidence into inherited semantic
regions, splits content when capacity is exceeded, and validates the resulting
package. PowerPoint generation reuses and duplicates source slides without
adding overlay text boxes. Excel formulas and fixed labels are preserved. Flat
PDFs without form fields are rebuilt on clean pages at the template's page size
because their original content has no safe editable slots.

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

The deployment build targets Cloudflare Workers through vinext. Standard
Cloudflare configuration lives in `wrangler.jsonc`: D1 is bound as `DB`, R2 is
bound as `FILES`, and generated SQL migrations are stored in `drizzle/`.

Before deploying, create a D1 database and R2 bucket in the target Cloudflare
account. Replace the placeholder D1 `database_id` in `wrangler.jsonc` with the
created database ID, and adjust the resource names if needed. Then deploy with:

```bash
npm run deploy
```

## Honest limitations

This release establishes the conversational, ingestion, grounding, and
persistence foundation. The next production phases are:

- a layout-aware PDF parser and multimodal inspection for source images and
  slide visuals;
- a legacy `.xls` compatibility adapter;
- hierarchical retrieval over large corpora and deterministic conflict
  normalization;
- full knowledge-version and source-priority-rule services;
- richer visual inference for image templates and non-fillable PDF artwork;
- native template editing for legacy `.xls` files (convert them to `.xlsx` today);
- authentication policy, enterprise tenancy, malware scanning, retention,
  observability, and provider cost estimation.

Those gaps are surfaced in extraction status and documentation rather than
represented as complete.
