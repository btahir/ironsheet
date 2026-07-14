// Local TypeDoc plugin: give every generated Markdown page a per-page
// `title:` frontmatter value so Fumadocs (which requires a title on every
// page) renders the reference without hand-editing. Works together with
// typedoc-plugin-frontmatter, which serializes `page.frontmatter` to YAML.
import { MarkdownPageEvent } from "typedoc-plugin-markdown";

export function load(app) {
  app.renderer.on(MarkdownPageEvent.BEGIN, (page) => {
    const project = page.project;
    const isRoot = page.model === project || page.url === app.renderer.router?.getFullUrl(project);

    const title = isRoot ? "API Reference" : (page.model?.name ?? "Reference");

    page.frontmatter = {
      title,
      ...(page.frontmatter ?? {})
    };
  });
}
