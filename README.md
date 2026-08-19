# PMI Agent — Decision-ready integration reporting

PMI Agent is a conversational Post-Merger Integration reporting workspace. It
combines project context, uploaded evidence, source coverage, and a configured
OpenAI or Anthropic model to produce traceable management answers and report
drafts.

This README is the complete public guide to the project. It does not rely on
supplementary project documentation.

## What the application does

PMI Agent is intended for Boards, Steering Committees, Integration Management
Offices, finance leaders, and workstream teams. A user can:

- keep project context and conversations separate;
- attach multiple evidence files to a message;
- see whether each file was extracted completely, partially, or not at all;
- ask questions across the available evidence;
- inspect extraction warnings and material cross-source conflicts;
- generate editable PPTX, XLSX, and DOCX reports or fixed PDF and HTML outputs;
- revise a previous report while retaining its version history; and
- download generated artifacts from the message that created them.

The project does **not** train or fine-tune a language model. The “agent” is the
complete application workflow: PMI-specific instructions, project and audience
context, source extraction, evidence controls, model synthesis, deterministic
rendering, persistence, and versioning.

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
- source coverage inspection, extraction warnings, and deterministic conflict
  visibility;
- D1-backed project, chat, message, source, and artifact metadata persistence;
- version-ready knowledge, authority-rule, draft, provenance, and model-usage
  schema;
- explicit PowerPoint intent and revision handling with a structured,
  model-owned storyline;
- optional `@filename` template selection for PPTX, XLSX, DOCX, PDF, and HTML;
- server-side report rendering, R2 storage, and per-message downloads; and
- a responsive consulting-chat interface.

## How a request is processed

1. The application combines the user request with the selected audience,
   project context, recent conversation, and every attached source.
2. Each source is extracted and assigned a coverage status and warnings.
3. Comparable evidence is reconciled. Unresolved disagreements remain visible;
   the application does not silently average or select a conflicting value.
4. For ordinary chat, the selected provider streams a grounded answer.
5. For a report request, the model returns a constrained semantic content model.
6. Deterministic TypeScript renderers turn that model into the requested file,
   validate the result, store it, and return a download link.
7. A revision can load the previous structured report and create a new version.

The model decides the synthesis, implications, storyline, and clearly labelled
recommendations. Application code controls source enumeration, extraction
status, conflict visibility, file construction, validation, and persistence.

## Supported source files

| Input | Extracted information | Coverage |
| --- | --- | --- |
| CSV | Headers and row records | Deterministic |
| XLSX | Sheets, rows, cell references, values, and formulas | Deterministic |
| DOCX | Paragraphs and basic transcript signals | Deterministic |
| PPTX | Slide text and speaker notes | Deterministic for extracted text |
| HTML | Semantic blocks | Deterministic |
| PDF | Conservative text-layer content | Partial or low-confidence when appropriate |
| PNG/JPEG | Dimensions and a pending-inspection marker | Visual meaning is not assumed |
| XLS | Compatibility warning | Legacy adapter not yet implemented |

An individual message accepts up to 20 files, with a maximum size of 25 MB per
file. Large extracted documents are bounded before being sent to the model; the
current release does not use hierarchical or vector retrieval.

## Generated reports

| Format | Best suited to | Result |
| --- | --- | --- |
| PPTX | Board and Steering Committee presentations | Editable slides |
| XLSX | Trackers, registers, KPI views, and auditable detail | Editable workbook |
| DOCX | Narrative reports, memos, and decision papers | Editable document |
| PDF | Fixed, distribution-ready reports | Non-editable document |
| HTML | Portable dashboards | Self-contained responsive file |

### Using a report template

Attach the template and the PMI evidence files in the same chat. Click the `@`
button on the template attachment, or type its exact `@filename`, and request a
new report. For example:

```text
Using @Executive_Status_Template.pptx, create the same report for our PMI deal.
```

The marked file controls the output structure and is excluded from deal
evidence; every other attachment remains an evidence source. Prompts without an
`@filename` use the standard artifact workflow.

Template generation is structure-first. The server inspects the selected file,
maps evidence into inherited semantic regions, splits content when capacity is
exceeded, and validates the resulting package. PowerPoint generation reuses and
duplicates source slides without adding overlay text boxes. Excel formulas and
fixed labels are preserved. Flat PDFs without form fields are rebuilt on clean
pages at the template page size because their original content has no safe
editable slots.

## Local setup

### Prerequisites

- Node.js 22.13 or later (Node.js 24 LTS is recommended for new installations)
- npm, included with Node.js
- Git when cloning or updating the repository
- internet access for dependency installation and model-provider requests
- an API key and an enabled model ID for at least one supported provider

### Install and run

From the repository root:

```bash
npm install
cp .env.example .env.local
npm run dev
```

Open the local address printed in the terminal, normally
`http://localhost:3000`. Keep the terminal running while using the application;
press `Ctrl+C` to stop it.

The development command uses Wrangler/Miniflare, including local D1 and R2
bindings. The core application tables are created on first use.

### Windows setup for non-developers

1. Install Node.js 24 LTS, Git for Windows, and Visual Studio Code. If they are
   managed by your organization, use the company software portal.
2. Restart Visual Studio Code after installation, open **Terminal > New
   Terminal**, and verify:

   ```powershell
   node --version
   npm --version
   git --version
   ```

3. Clone the supplied repository URL, or extract the supplied ZIP and open the
   folder containing `package.json` in Visual Studio Code.
4. In the PowerShell terminal, run:

   ```powershell
   npm install
   Copy-Item .env.example .env.local
   npm run dev
   ```

5. If PowerShell blocks `npm.ps1`, use `npm.cmd install` and
   `npm.cmd run dev`. This avoids changing the machine's execution policy.
6. Open `http://localhost:3000`, select a configured model, and test with a
   small non-confidential file.

The first setup normally takes longer because npm downloads the project
dependencies. Do not run `npm audit fix --force`; escalate reported dependency
issues to the project maintainer instead.

## Provider configuration

Edit `.env.local` and configure at least one provider key plus at least one
matching model ID:

```dotenv
OPENAI_API_KEY=
ANTHROPIC_API_KEY=

OPENAI_GPT55_MODEL=
OPENAI_GPT56_MODEL=
OPENAI_MINI_MODEL=
ANTHROPIC_SONNET_MODEL=
ANTHROPIC_OPUS_MODEL=
```

The names shown in the interface are display labels, not guaranteed provider
model IDs. Copy an exact, currently available model ID from the provider account
or official documentation. Configure only models the account can use and leave
the other entries empty.

Provider keys and actual model IDs remain server-side. A model is available in
the selector only when both its API key and model ID are present. Restart the
development server after changing `.env.local`.

### Secret handling

- Never place real keys in `.env.example`, source code, screenshots, chat, or a
  Git commit.
- Keep `.env.local` private; it is excluded from source control.
- Prefer individual or service-specific keys and apply provider usage limits.
- Revoke and replace a key immediately if it is exposed.
- Follow organizational policy before sending confidential business data to a
  model provider.

## First-use check

1. Create or open a project or standalone chat and choose the intended audience.
2. Select a model without a configuration warning.
3. Send a simple message without confidential information.
4. Upload a small approved file and confirm its extraction status is shown.
5. Request an answer or report, review the evidence notes and warnings, and
   download the generated artifact.

## Validation

Run the repository checks from the project root:

```bash
npm test
npm run lint
npm run db:generate
```

`npm test` performs a production build before running the Node test suite.
`npm run db:generate` creates SQL migrations after intentional schema changes;
review generated migration files before committing them.

## Deployment to Cloudflare

The production build targets Cloudflare Workers through vinext. Configuration
lives in `wrangler.jsonc`: D1 is bound as `DB`, R2 is bound as `FILES`, and SQL
migrations are stored in `drizzle/`.

Before deploying:

1. Create a D1 database and an R2 bucket in the target Cloudflare account.
2. Replace the placeholder D1 `database_id` in `wrangler.jsonc` with the real
   database ID and adjust resource names if necessary.
3. Configure provider secrets and model IDs in the deployment environment.
4. Apply the repository migrations to the target D1 database according to the
   organization's release process.
5. Deploy:

   ```bash
   npm run deploy
   ```

The application is not hosted by a model provider and `npm run dev` does not
publish it to the internet.

## Architecture and source map

PMI Agent is a TypeScript/React application using a Next-compatible App Router,
vinext, and a Cloudflare Worker-compatible ESM runtime. D1/SQLite stores
structured state through the `DB` binding; R2-compatible object storage stores
original uploads and generated artifacts through `FILES`.

| Concern | Primary implementation |
| --- | --- |
| Chat and artifact orchestration | `app/api/chat/route.ts` |
| PMI runtime instructions | `app/lib/pmi-prompt.ts` |
| Model registry and providers | `app/lib/models.ts`, `app/lib/llm/` |
| Extraction and source coverage | `app/api/extract/route.ts`, `app/lib/ingestion/` |
| Evidence reconciliation | `app/lib/evidence.ts`, `app/lib/conflict-guard.ts` |
| Artifact planning | `app/lib/artifact-intent.ts`, `app/lib/artifact.ts` |
| PPTX generation | `app/lib/presentation.ts` |
| XLSX, DOCX, PDF, and HTML generation | `app/lib/renderers/` |
| Persistence and runtime bindings | `app/lib/persistence.ts`, `app/lib/runtime-bindings.ts` |
| Relational schema and migrations | `db/`, `drizzle/` |
| Cloudflare configuration | `wrangler.jsonc`, `vite.config.ts` |

Runtime report generation is implemented in TypeScript/JavaScript; it does not
depend on Python. Structured state includes projects, chats, messages, sources,
extracted segments, source-priority rules, knowledge versions, drafts, artifact
versions, provenance, and model-usage records.

## Troubleshooting

| Symptom | What to check |
| --- | --- |
| `node`, `npm`, or `git` is not recognized | Restart the terminal, then verify the tool was installed and added to `PATH`. |
| PowerShell says `npm.ps1` cannot be loaded | Use `npm.cmd` instead of changing the execution policy. |
| A model says “Configure …” | Add both the provider key and matching model ID to `.env.local`, then restart. |
| Authentication error | Confirm the key belongs to the selected provider and has not expired or been revoked. |
| Quota, credit, or billing error | Review the provider account's billing, usage, and limits. |
| Model unavailable or not found | Use an exact model ID enabled for the account. |
| Port 3000 is already in use | Stop the other local server or use the alternate address printed by the dev command. |
| Upload is only partially extracted | Review the warning and validate manually or provide a clearer source format. |
| `npm install` reports package funding | No action is required. |

For a Git installation, update only when the project owner announces a release:

```bash
git pull
npm install
npm run dev
```

If Git reports local changes or a merge conflict, stop and ask the maintainer
before continuing. Do not discard files or use force commands.

## Responsible-use boundaries

PMI Agent produces drafts for professional review; it is not an autonomous
source of corporate truth. Users must validate consequential financial values,
dates, owners, risks, decisions, calculations, and recommendations.

The principal current gaps are:

- layout-aware PDF parsing and multimodal inspection for images and slide
  visuals;
- legacy `.xls` extraction and native template editing (convert to `.xlsx`);
- hierarchical retrieval for large corpora and richer deterministic conflict
  normalization;
- complete knowledge-version and source-priority-rule services;
- richer visual inference for image templates and non-fillable PDF artwork; and
- production controls for authentication, authorization, enterprise tenancy,
  malware scanning, retention, observability, provider cost estimation, and
  organizational secret management.

These limitations are surfaced through extraction status and warnings rather
than represented as complete coverage.
