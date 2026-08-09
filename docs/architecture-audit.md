# PMI Reporting Agent — Architecture Audit and Delivery Plan

## Repository audit

The workspace was greenfield at the start of implementation. It contained only
`prompt.txt`; there was no application, source-control metadata, test suite, or
existing architecture to preserve. The application therefore uses the bundled
Sites full-stack starter as its deployment surface.

## Chosen architecture

- **Web application:** React 19, TypeScript, Next-compatible App Router, vinext.
- **Runtime:** Cloudflare Worker-compatible ESM.
- **Structured persistence:** D1/SQLite, declared as `DB`.
- **File storage:** R2, declared as `FILES`.
- **AI integration:** provider-neutral `LLMProvider` interface with OpenAI and
  Anthropic adapters. Display names are separate from provider model IDs.
- **Knowledge model:** project, chat, message, source, extracted segment,
  source-priority rule, report draft, and draft version records.
- **Grounding:** every response plan starts with the complete applicable source
  manifest. Sources are marked used, not used with a reason, or pending
  extraction; a response cannot claim complete coverage while an applicable
  source is unprocessed.
- **UI:** conversational workspace with projects and standalone chats, model
  selection, project context, multi-file messages, source inspection, streamed
  responses, and editable report drafts.

## Capability map before implementation

| Specification phase | Existing capability | Gap at audit |
| --- | --- | --- |
| Conversational core | None | Complete phase missing |
| Reliable ingestion | None | Complete phase missing |
| Knowledge and retrieval | None | Complete phase missing |
| PMI consultant reasoning | None | Complete phase missing |
| Editable drafts | None | Complete phase missing |
| Professional exports | None | Complete phase missing |
| Quality hardening | Starter render test only | Product tests missing |

## Staged delivery

1. Deliver the conversational product shell and durable domain model.
2. Add deterministic extraction adapters with preserved provenance and explicit
   warnings for anything not extractable in the runtime.
3. Build source-manifest retrieval, source coverage checks, authority rules,
   conflict/gap representation, and versioned knowledge.
4. Connect provider adapters to a PMI-specific synthesis and validation
   pipeline.
5. Persist report drafts and conversational revisions.
6. Add template-aware PPTX, XLSX, DOCX, and PDF renderers with visual QA.
7. Expand fixtures, leakage tests, security controls, observability, and cost
   reporting.

## PowerPoint generation update

PowerPoint intent now branches before the normal text stream. The model produces
a validated presentation model; a server-side renderer creates the PPTX; R2
stores the binary; and D1 associates artifact metadata and the complete model
with the assistant message. Follow-up slide revisions load that model, preserve
unaffected slides, and write a new immutable artifact version. Ordinary status
questions continue through the text chat path.

This document records the mandated pre-change audit. It is intentionally candid:
features are only marked implemented when their production path exists and is
covered by validation.
