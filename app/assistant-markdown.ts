import { createElement } from "react";
import ReactMarkdown, { type Components } from "react-markdown";
import remarkGfm from "remark-gfm";

const markdownComponents: Components = {
  table({ children, node, ...props }) {
    void node;
    return createElement(
      "div",
      { className: "markdown-table-wrap" },
      createElement("table", props, children),
    );
  },
  a({ children, node, ...props }) {
    void node;
    return createElement(
      "a",
      { ...props, target: "_blank", rel: "noopener noreferrer" },
      children,
    );
  },
};

export function AssistantMarkdown({ content }: { content: string }) {
  return createElement(
    ReactMarkdown,
    {
      remarkPlugins: [remarkGfm],
      skipHtml: true,
      components: markdownComponents,
    },
    content,
  );
}
