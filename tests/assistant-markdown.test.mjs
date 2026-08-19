import assert from "node:assert/strict";
import test from "node:test";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { AssistantMarkdown } from "../app/assistant-markdown.ts";

const acceptanceMarkdown = `# Project Aurora — Board Summary

## Executive message

**Management-confirmed position:** Overall integration progress is **78%**.

*Sources: weekly_update.pptx.*

- **ERP cutover:** At risk
  1. Confirm owner

> Management attention required.

| Measure | Amount | % of target |
|---|---:|---:|
| Confirmed target | **€37.0m** | 100.0% |
| Forecast | **€28.7m** | 77.6% |

- Target: \`€18.0m + €6.0m = €24.0m\`

[Evidence](https://example.com)

---

\`\`\`text
fenced code
\`\`\`

<script>globalThis.compromised = true</script>`;

test("renders assistant Markdown as safe semantic HTML with GFM tables", () => {
  const html = renderToStaticMarkup(
    createElement(AssistantMarkdown, { content: acceptanceMarkdown }),
  );

  assert.match(html, /<h1>Project Aurora — Board Summary<\/h1>/);
  assert.match(html, /<h2>Executive message<\/h2>/);
  assert.match(html, /<strong>Management-confirmed position:<\/strong>/);
  assert.match(html, /<em>Sources: weekly_update\.pptx\.<\/em>/);
  assert.match(html, /<ul>/);
  assert.match(html, /<ol>/);
  assert.match(html, /<blockquote>/);
  assert.match(html, /<div class="markdown-table-wrap"><table>/);
  assert.match(html, /style="text-align:right"/);
  assert.match(html, /<td style="text-align:right"><strong>€37\.0m<\/strong><\/td>/);
  assert.match(html, /<code>€18\.0m \+ €6\.0m = €24\.0m<\/code>/);
  assert.match(html, /<pre><code class="language-text">fenced code\n<\/code><\/pre>/);
  assert.match(html, /href="https:\/\/example\.com" target="_blank" rel="noopener noreferrer"/);
  assert.match(html, /<hr\/>/);
  assert.doesNotMatch(html, /<script|compromised/);
  assert.doesNotMatch(html, /\|---|\*\*|```/);
});
