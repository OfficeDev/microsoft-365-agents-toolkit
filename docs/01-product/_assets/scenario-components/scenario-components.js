const optionTagName = "vscode-option";

const pickerDefinitions = new Map([
  ["vscode-single-select", { multiple: false, file: false, confirm: false }],
  ["vscode-multi-select", { multiple: true, file: false, confirm: true }],
  ["vscode-file-select", { multiple: false, file: true, confirm: false }],
]);

const iconFileNames = new Set([
  "teamsfx-agent",
  "teamsfx-custom-copilot",
  "teamsfx-graph-connector",
  "microsoft365-agents-toolkit-teams",
  "microsoft365-agents-office",
  "question",
  "file",
  "new-file",
  "folder",
]);

const iconSvgCache = new Map();

class VsCodePicker extends HTMLElement {
  connectedCallback() {
    const definition = pickerDefinitions.get(this.localName);
    if (!definition) return;

    const title = this.getAttribute("title") || "Select an option";
    const placeholder = this.getAttribute("placeholder") || "";
    const selectedCount = Array.from(this.querySelectorAll(`${optionTagName}[selected]`)).length;
    const selectedLabel = this.getAttribute("selected-label") || `${selectedCount} Selected`;
    const confirmLabel = this.getAttribute("confirm-label") || "OK";
    const options = Array.from(this.querySelectorAll(optionTagName)).map((option, index, all) => {
      const meta = option.getAttribute("meta") || "";
      const previousMeta = index > 0 ? all[index - 1].getAttribute("meta") || "" : "";
      const isNewGroup = !!meta && meta !== previousMeta;
      return {
        label: option.getAttribute("label") || "Untitled option",
        description: option.getAttribute("description") || "",
        detail: option.getAttribute("detail") || "",
        meta,
        showMeta: isNewGroup,
        dividerAbove: isNewGroup && index > 0,
        icon: option.getAttribute("icon") || "",
        selected: option.hasAttribute("selected"),
        active: option.hasAttribute("active") || option.hasAttribute("selected"),
      };
    });

    this.innerHTML = `
      <div class="vscode-command" role="img" aria-label="${escapeHtml(title)}">
        <div class="vscode-command__titlebar">
          <span class="vscode-command__back" aria-hidden="true"></span>
          <h4>${escapeHtml(title)}</h4>
        </div>
        <div class="vscode-command__input-row">
          <input class="vscode-command__input" placeholder="${escapeHtml(placeholder)}" aria-label="${escapeHtml(placeholder || title)}" readonly>
          ${definition.confirm ? `<span class="vscode-command__pill">${escapeHtml(selectedLabel)}</span><span class="vscode-command__button">${escapeHtml(confirmLabel)}</span>` : ""}
        </div>
        <ul class="vscode-options">
          ${options.map((option) => renderOption(option, definition)).join("")}
        </ul>
      </div>`;

    hydrateIcons(this);
  }
}

class VsCodeInputBox extends HTMLElement {
  connectedCallback() {
    const title = this.getAttribute("title") || "Input";
    const placeholder = this.getAttribute("placeholder") || "";
    const value = this.getAttribute("value") || "";
    const hint = this.getAttribute("hint") || "Press 'Enter' to confirm your input or 'Escape' to cancel";
    const error = this.getAttribute("error") || "";
    const hasError = error.length > 0;
    const footer = hasError
      ? `<div class="vscode-command__error" role="alert">${escapeHtml(error)}</div>`
      : `<p class="vscode-command__hint">${escapeHtml(hint)}</p>`;

    this.innerHTML = `
      <div class="vscode-command vscode-command--input${hasError ? " vscode-command--has-error" : ""}" role="img" aria-label="${escapeHtml(title)}">
        <div class="vscode-command__titlebar">
          <span class="vscode-command__back" aria-hidden="true"></span>
          <h4>${escapeHtml(title)}</h4>
        </div>
        <div class="vscode-command__input-row vscode-command__input-row--solo">
          <input class="vscode-command__input" value="${escapeHtml(value)}" placeholder="${escapeHtml(placeholder)}" aria-label="${escapeHtml(title)}" readonly>
        </div>
        ${footer}
      </div>`;
  }
}

class VsCodeCodeLensFile extends HTMLElement {
  connectedCallback() {
    const title = this.getAttribute("title") || ".vscode/mcp.json";
    const lens = this.getAttribute("lens") || "ATK: Fetch action from MCP | Start | tools | prompts | More...";
    const server = this.getAttribute("server") || "apigithubc";
    const type = this.getAttribute("type") || "http";
    const url = this.getAttribute("url") || "https://api.githubcopilot.com/mcp/";

    this.innerHTML = `
      <div class="vscode-code-file" role="img" aria-label="${escapeHtml(title)} with CodeLens">
        <div class="vscode-code-file__bar">${escapeHtml(title)}</div>
        <div class="vscode-code-file__body">
          ${renderCodeRow(1, `<span class="vscode-code-token--punctuation">{</span>`)}
          ${renderCodeRow(2, `<span class="vscode-code-indent-1"><span class="vscode-code-token--string">&quot;servers&quot;</span>: <span class="vscode-code-token--punctuation">{</span></span>`)}
          ${renderCodeRow(3, `<span class="vscode-code-codelens">${escapeHtml(lens)}</span><span class="vscode-code-indent-2"><span class="vscode-code-token--string">&quot;${escapeHtml(server)}&quot;</span>: <span class="vscode-code-token--punctuation">{</span></span>`)}
          ${renderCodeRow(4, `<span class="vscode-code-indent-3"><span class="vscode-code-token--string">&quot;type&quot;</span>: <span class="vscode-code-token--string">&quot;${escapeHtml(type)}&quot;</span>,</span>`)}
          ${renderCodeRow(5, `<span class="vscode-code-indent-3"><span class="vscode-code-token--string">&quot;url&quot;</span>: <span class="vscode-code-token--string">&quot;${escapeHtml(url)}&quot;</span></span>`, true)}
          ${renderCodeRow(6, `<span class="vscode-code-indent-2"><span class="vscode-code-token--punctuation">}</span></span>`)}
          ${renderCodeRow(7, `<span class="vscode-code-indent-1"><span class="vscode-code-token--punctuation">}</span></span>`)}
          ${renderCodeRow(8, `<span class="vscode-code-token--punctuation">}</span>`)}
        </div>
      </div>`;
  }
}

class VsCodeNotification extends HTMLElement {
  connectedCallback() {
    const title = this.getAttribute("title") || "Notification";
    const message = this.getAttribute("message") || "";
    const detail = this.getAttribute("detail") || "";
    const action = this.getAttribute("action") || "";
    const source = this.getAttribute("source") || "";
    const rawSeverity = (this.getAttribute("severity") || "info").toLowerCase();
    const severity = ["info", "warning", "error"].includes(rawSeverity) ? rawSeverity : "info";
    const iconChar = severity === "error" ? "\u2715" : severity === "warning" ? "!" : "i";
    const role = severity === "error" ? "alert" : "status";

    this.innerHTML = `
      <div class="vscode-notification vscode-notification--${severity}" role="${role}" aria-label="${escapeHtml(title)}">
        <div class="vscode-notification__icon" aria-hidden="true">${iconChar}</div>
        <div class="vscode-notification__body">
          <p class="vscode-notification__message">${escapeHtml(message)}</p>
          ${detail ? `<p class="vscode-notification__detail">${escapeHtml(detail)}</p>` : ""}
          ${source || action ? `
            <div class="vscode-notification__footer">
              ${source ? `<span class="vscode-notification__source">Source: ${escapeHtml(source)}</span>` : `<span></span>`}
              ${action ? `<span class="vscode-notification__action">${escapeHtml(action)}</span>` : ""}
            </div>` : ""}
        </div>
        <div class="vscode-notification__controls" aria-hidden="true">
          <span class="vscode-notification__control" title="Configure notification">&#9881;</span>
          <span class="vscode-notification__control" title="Clear notification">&times;</span>
        </div>
      </div>`;
  }
}

class VsCodeModalNotification extends HTMLElement {
  connectedCallback() {
    const windowTitle = this.getAttribute("window-title") || "Visual Studio Code";
    const message = this.getAttribute("message") || "";
    const detail = this.getAttribute("detail") || "";
    const rawSeverity = (this.getAttribute("severity") || "info").toLowerCase();
    const severity = ["info", "warning", "error"].includes(rawSeverity) ? rawSeverity : "info";
    const role = severity === "error" ? "alertdialog" : "dialog";
    const buttonsAttr = this.getAttribute("buttons") || "OK";
    const buttons = buttonsAttr
      .split("|")
      .map((label) => label.trim())
      .filter(Boolean)
      .map((label, index) => ({ label, primary: index === 0 }));
    const ariaLabel = `${windowTitle}: ${message || detail || "Modal notification"}`;
    const iconSvg =
      severity === "warning"
        ? `<svg viewBox="0 0 24 24" width="32" height="32" aria-hidden="true" focusable="false"><path d="M12 2.5 1.5 21.5h21L12 2.5z" fill="#f5c324" stroke="#a37a00" stroke-width="0.6" stroke-linejoin="round"/><rect x="11.05" y="9.5" width="1.9" height="6.6" rx="0.4" fill="#1a1a1a"/><rect x="11.05" y="17.2" width="1.9" height="1.9" rx="0.4" fill="#1a1a1a"/></svg>`
        : severity === "error"
        ? `<svg viewBox="0 0 24 24" width="32" height="32" aria-hidden="true" focusable="false"><circle cx="12" cy="12" r="10" fill="#e51400"/><path d="M8 8l8 8M16 8l-8 8" stroke="#ffffff" stroke-width="2" stroke-linecap="round"/></svg>`
        : `<svg viewBox="0 0 24 24" width="32" height="32" aria-hidden="true" focusable="false"><circle cx="12" cy="12" r="10" fill="#0078d4"/><path d="M12 7.5v.1m0 3v6" stroke="#ffffff" stroke-width="2" stroke-linecap="round"/></svg>`;

    this.innerHTML = `
      <div class="vscode-modal-notification vscode-modal-notification--${severity}" role="${role}" aria-modal="true" aria-label="${escapeHtml(ariaLabel)}">
        <div class="vscode-modal-notification__dialog">
          <div class="vscode-modal-notification__titlebar" aria-hidden="true">
            <span class="vscode-modal-notification__window-title">${escapeHtml(windowTitle)}</span>
            <span class="vscode-modal-notification__close" title="Close">&times;</span>
          </div>
          <div class="vscode-modal-notification__body">
            <div class="vscode-modal-notification__icon" aria-hidden="true">${iconSvg}</div>
            <div class="vscode-modal-notification__content">
              ${message ? `<p class="vscode-modal-notification__message">${escapeHtml(message)}</p>` : ""}
              ${detail ? `<p class="vscode-modal-notification__detail">${escapeHtml(detail)}</p>` : ""}
            </div>
          </div>
          <div class="vscode-modal-notification__actions">
            ${buttons
              .map(
                (btn) =>
                  `<button type="button" class="vscode-modal-notification__button${btn.primary ? " vscode-modal-notification__button--primary" : ""}" tabindex="-1">${escapeHtml(btn.label)}</button>`
              )
              .join("")}
          </div>
        </div>
      </div>`;
  }
}

class ScenarioMermaidFlow extends HTMLElement {
  connectedCallback() {
    const source = this.getAttribute("src") || "";
    const title = this.getAttribute("title") || "Flow reference";
    const flowIndex = parseFlowIndex(this.getAttribute("flow-index"));

    this.innerHTML = `
      <div class="scenario-mermaid" aria-label="${escapeHtml(title)}">
        <div class="scenario-mermaid__head">
          <h3>${escapeHtml(title)}</h3>
          <p>${source ? escapeHtml(`${source} #${flowIndex + 1}`) : "No Markdown source configured."}</p>
        </div>
        <div class="scenario-mermaid__body" aria-live="polite">Loading flow...</div>
      </div>`;

    if (!source) {
      this.querySelector(".scenario-mermaid__body").textContent = "No Markdown source configured.";
      return;
    }

    renderMarkdownMermaidFlow(this, source, flowIndex);
  }
}

class ScenarioMarkdownSection extends HTMLElement {
  connectedCallback() {
    const source = this.getAttribute("src") || "";
    const heading = this.getAttribute("heading") || "";
    const level = clampHeadingLevel(this.getAttribute("level"), 3);
    const title = this.getAttribute("title") || heading || "Markdown section";

    this.innerHTML = `
      <div class="scenario-markdown" aria-label="${escapeHtml(title)}">
        <div class="scenario-markdown__head">
          <h3>${escapeHtml(title)}</h3>
          <p>${source && heading ? escapeHtml(`${source} \u00a7 ${heading}`) : "No Markdown source configured."}</p>
        </div>
        <div class="scenario-markdown__body" aria-live="polite">Loading section...</div>
      </div>`;

    if (!source || !heading) {
      this.querySelector(".scenario-markdown__body").textContent = "Missing src or heading attribute.";
      return;
    }

    renderMarkdownSection(this, source, heading, level);
  }
}

function renderCodeRow(lineNumber, content, active = false) {
  const activeClass = active ? " vscode-code-row--active" : "";
  return `<div class="vscode-code-row${activeClass}"><span class="vscode-code-line">${lineNumber}</span><span class="vscode-code-text">${content}</span></div>`;
}

function renderOption(option, definition) {
  const marker = getOptionMarker(option, definition);
  const stateClasses = [option.active ? "vscode-option--active" : "", option.selected ? "vscode-option--checked" : ""].filter(Boolean).join(" ");
  const markerClass = marker ? "" : " vscode-option--no-marker";
  const description = option.description ? `<span class="vscode-option__description">${escapeHtml(option.description)}</span>` : "";
  const detail = option.detail ? `<span class="vscode-option__path">${escapeHtml(option.detail)}</span>` : "";
  const meta = option.showMeta ? `<span class="vscode-option__meta">${escapeHtml(option.meta)}</span>` : "";
  const dividerClass = option.dividerAbove ? " vscode-option--group-start" : "";

  return `
    <li class="vscode-option${markerClass}${dividerClass} ${stateClasses}" aria-label="${escapeHtml(option.label)}">
      ${marker}
      <span>
        <span class="vscode-option__heading">
          <span class="vscode-option__label">${escapeHtml(option.label)}</span>
          ${meta}
        </span>
        ${description}${detail}
      </span>
    </li>`;
}

function getOptionMarker(option, definition) {
  if (definition.file) return renderFileIcon(option);
  if (definition.multiple) return `<span class="vscode-option__mark" aria-hidden="true"></span>`;
  if (!option.icon) return "";

  const icon = normalizeIconName(option.icon);
  return renderIconSlot("vscode-option-icon", icon);
}

function normalizeIconName(value) {
  const iconAliases = {
    declarative: "teamsfx-agent",
    engine: "teamsfx-custom-copilot",
    connector: "teamsfx-graph-connector",
    teams: "microsoft365-agents-toolkit-teams",
    office: "microsoft365-agents-office",
    copilot: "question",
    file: "default",
  };
  const icon = value.trim().replace(/^\$\((.+)\)$/, "$1");
  return iconAliases[icon] || icon || "default";
}

function renderFileIcon(option) {
  const icon = normalizeFileIconName(option.icon);
  return renderIconSlot("vscode-file-icon", icon);
}

function normalizeFileIconName(value) {
  const iconAliases = {
    new: "new-file",
    file: "file",
  };
  const icon = value.trim().replace(/^\$\((.+)\)$/, "$1");
  return iconAliases[icon] || icon || "file";
}

function renderIconSlot(className, icon) {
  const iconName = iconFileNames.has(icon) ? icon : "file";
  return `<span class="${className} ${className}--${escapeClassName(iconName)}" data-vscode-icon="${escapeHtml(iconName)}" aria-hidden="true"></span>`;
}

function hydrateIcons(root) {
  const iconSlots = Array.from(root.querySelectorAll("[data-vscode-icon]"));
  for (const iconSlot of iconSlots) {
    const iconName = iconSlot.getAttribute("data-vscode-icon");
    if (!iconName || !iconFileNames.has(iconName)) continue;
    loadIconSvg(iconName).then((svg) => {
      iconSlot.innerHTML = svg;
    });
  }
}

function loadIconSvg(iconName) {
  if (!iconSvgCache.has(iconName)) {
    const iconUrl = new URL(`./icons/${iconName}.svg`, import.meta.url).href;
    iconSvgCache.set(
      iconName,
      fetch(iconUrl)
        .then((response) => (response.ok ? response.text() : ""))
        .then((svg) => svg.replace("<svg ", '<svg class="vscode-inline-icon" focusable="false" aria-hidden="true" '))
        .catch(() => "")
    );
  }
  return iconSvgCache.get(iconName);
}

async function renderMarkdownMermaidFlow(root, source, flowIndex = 0) {
  const body = root.querySelector(".scenario-mermaid__body");
  try {
    const response = await fetch(source, { cache: "no-store" });
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
    const markdown = await response.text();
    const mermaid = extractMermaidFlow(markdown, flowIndex);
    if (!mermaid) throw new Error("No Mermaid code block found in Markdown source.");

    body.innerHTML = `<div class="mermaid">${escapeHtml(mermaid)}</div>`;
    const mermaidApi = await import("https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs");
    mermaidApi.default.initialize({ startOnLoad: false, securityLevel: "strict", theme: "default" });
    await mermaidApi.default.run({ nodes: [body.querySelector(".mermaid")] });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    body.innerHTML = `<p class="scenario-mermaid__warning">Flow preview unavailable: ${escapeHtml(message)}</p>`;
  }
}

function extractMermaidFlow(markdown, flowIndex = 0) {
  const matches = Array.from(markdown.replace(/\r\n/g, "\n").matchAll(/```mermaid\s*\n([\s\S]*?)\n```/gi));
  return matches[flowIndex]?.[1]?.trim() || "";
}

async function renderMarkdownSection(root, source, heading, level) {
  const body = root.querySelector(".scenario-markdown__body");
  try {
    const response = await fetch(source, { cache: "no-store" });
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
    const markdown = await response.text();
    const block = extractMarkdownSection(markdown, heading, level);
    if (!block) throw new Error(`Heading not found: ${heading}`);
    body.innerHTML = renderMarkdownBlocks(block);

    const mermaidNodes = Array.from(body.querySelectorAll(".mermaid"));
    if (mermaidNodes.length > 0) {
      const mermaidApi = await import("https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs");
      mermaidApi.default.initialize({ startOnLoad: false, securityLevel: "strict", theme: "default" });
      await mermaidApi.default.run({ nodes: mermaidNodes });
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    body.innerHTML = `<p class="scenario-markdown__warning">Section preview unavailable: ${escapeHtml(message)}</p>`;
  }
}

function extractMarkdownSection(markdown, heading, level) {
  const lines = markdown.replace(/\r\n/g, "\n").split("\n");
  const startNeedle = `${"#".repeat(level)} ${heading}`.trim();
  const stopRegex = new RegExp(`^#{1,${level}}\\s`);
  let startIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].trim() === startNeedle) {
      startIdx = i + 1;
      break;
    }
  }
  if (startIdx === -1) return "";
  let endIdx = lines.length;
  for (let i = startIdx; i < lines.length; i++) {
    if (stopRegex.test(lines[i])) {
      endIdx = i;
      break;
    }
  }
  return lines.slice(startIdx, endIdx).join("\n").trim();
}

function renderMarkdownBlocks(markdown) {
  const lines = markdown.split("\n");
  const html = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (/^```/.test(line)) {
      const fenceMatch = line.match(/^```\s*([\w-]+)?/);
      const lang = (fenceMatch?.[1] || "").toLowerCase();
      const codeLines = [];
      i++;
      while (i < lines.length && !/^```\s*$/.test(lines[i])) {
        codeLines.push(lines[i]);
        i++;
      }
      i++;
      const code = codeLines.join("\n");
      if (lang === "mermaid") {
        html.push(`<div class="mermaid">${escapeHtml(code)}</div>`);
      } else {
        const classAttr = lang ? ` class="language-${escapeClassName(lang)}"` : "";
        html.push(`<pre><code${classAttr}>${escapeHtml(code)}</code></pre>`);
      }
      continue;
    }
    if (line.trim() === "") {
      i++;
      continue;
    }
    if (/^####\s+/.test(line)) {
      html.push(`<h4>${renderMarkdownInline(line.replace(/^####\s+/, ""))}</h4>`);
      i++;
      continue;
    }
    if (/^[-*]\s+/.test(line)) {
      const items = [];
      while (i < lines.length && /^[-*]\s+/.test(lines[i])) {
        items.push(lines[i].replace(/^[-*]\s+/, ""));
        i++;
      }
      html.push(`<ul>${items.map((item) => `<li>${renderMarkdownInline(item)}</li>`).join("")}</ul>`);
      continue;
    }
    const paraLines = [line];
    i++;
    while (
      i < lines.length &&
      lines[i].trim() !== "" &&
      !/^[-*]\s+/.test(lines[i]) &&
      !/^```/.test(lines[i]) &&
      !/^#{1,6}\s/.test(lines[i])
    ) {
      paraLines.push(lines[i]);
      i++;
    }
    html.push(`<p>${renderMarkdownInline(paraLines.join(" "))}</p>`);
  }
  return html.join("\n");
}

function renderMarkdownInline(text) {
  let out = escapeHtml(text);
  out = out.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, (_m, label, href) => `<a href="${safeHref(href)}">${label}</a>`);
  out = out.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
  out = out.replace(/`([^`]+)`/g, "<code>$1</code>");
  return out;
}

function safeHref(href) {
  const trimmed = (href || "").trim();
  if (!trimmed) return "#";
  if (/^(#|\/(?!\/)|\.{0,2}\/)/.test(trimmed)) return trimmed;
  if (/^(https?:|mailto:)/i.test(trimmed)) return trimmed;
  return "#";
}

function clampHeadingLevel(value, fallback) {
  const parsed = Number.parseInt(value || `${fallback}`, 10);
  if (!Number.isFinite(parsed)) return fallback;
  if (parsed < 1) return 1;
  if (parsed > 6) return 6;
  return parsed;
}

function parseFlowIndex(value) {
  if (value === null) return 0;
  const parsed = Number.parseInt(value || "0", 10);
  if (!Number.isFinite(parsed) || parsed <= 1) return 0;
  return parsed - 1;
}

function escapeClassName(value) {
  return value.toLowerCase().replace(/[^a-z0-9_-]/g, "") || "default";
}

function escapeHtml(value) {
  return value.replace(/[&<>"]/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[character]));
}

for (const tagName of pickerDefinitions.keys()) {
  customElements.define(tagName, class extends VsCodePicker {});
}

customElements.define("vscode-input-box", VsCodeInputBox);
customElements.define("vscode-codelens-file", VsCodeCodeLensFile);
customElements.define("vscode-notification", VsCodeNotification);
customElements.define("vscode-modal-notification", VsCodeModalNotification);
customElements.define("scenario-mermaid-flow", ScenarioMermaidFlow);
customElements.define("scenario-markdown-section", ScenarioMarkdownSection);

function buildScenarioToc() {
  const main = document.querySelector("main");
  if (!main) return;
  const sections = Array.from(main.querySelectorAll("section[aria-labelledby]"));
  const items = sections
    .map((section) => {
      const headingId = section.getAttribute("aria-labelledby");
      const heading = headingId ? document.getElementById(headingId) : null;
      if (!heading) return null;
      if (!section.id) section.id = `${headingId}-section`;
      return { section, heading, id: section.id, text: heading.textContent.trim() };
    })
    .filter(Boolean);
  if (items.length < 2) return;

  const aside = document.createElement("aside");
  aside.className = "scenario-toc";
  aside.setAttribute("aria-label", "On this page");
  const details = document.createElement("details");
  details.className = "scenario-toc__details";
  const summary = document.createElement("summary");
  summary.className = "scenario-toc__title";
  summary.textContent = "On this page";
  const list = document.createElement("ol");
  list.className = "scenario-toc__list";
  for (const item of items) {
    const li = document.createElement("li");
    const link = document.createElement("a");
    link.href = `#${encodeURIComponent(item.id)}`;
    link.textContent = item.text;
    link.dataset.tocTarget = item.id;
    li.appendChild(link);
    list.appendChild(li);
  }
  details.appendChild(summary);
  details.appendChild(list);
  aside.appendChild(details);
  document.body.classList.add("has-scenario-toc");
  document.body.insertBefore(aside, document.body.firstChild);

  const wideQuery = window.matchMedia("(min-width: 1320px)");
  const syncOpen = () => {
    details.open = wideQuery.matches;
  };
  syncOpen();
  if (typeof wideQuery.addEventListener === "function") {
    wideQuery.addEventListener("change", syncOpen);
  } else if (typeof wideQuery.addListener === "function") {
    wideQuery.addListener(syncOpen);
  }
  list.addEventListener("click", (event) => {
    if (!wideQuery.matches && event.target.closest("a")) {
      details.open = false;
    }
  });

  const links = new Map(items.map((item) => [item.id, aside.querySelector(`a[data-toc-target="${item.id}"]`)]));
  const setActive = (id) => {
    for (const [linkId, link] of links) link.classList.toggle("is-active", linkId === id);
  };
  if ("IntersectionObserver" in window) {
    const observer = new IntersectionObserver(
      (entries) => {
        const visible = entries
          .filter((entry) => entry.isIntersecting)
          .sort((a, b) => a.boundingClientRect.top - b.boundingClientRect.top);
        if (visible.length > 0) setActive(visible[0].target.id);
      },
      { rootMargin: "-30% 0px -60% 0px", threshold: 0 }
    );
    for (const item of items) observer.observe(item.section);
  }
  setActive(items[0].id);
}

function makeScenarioFlowsCollapsible() {
  const sections = document.querySelectorAll("section.scenario-flow");
  for (const section of sections) {
    const head = section.querySelector(":scope > .section-head");
    if (!head) continue;
    const heading = head.querySelector(":scope > h2");
    if (!heading) continue;
    if (section.classList.contains("scenario-flow--collapsible")) continue;

    section.classList.add("scenario-flow--collapsible");
    head.setAttribute("role", "button");
    head.setAttribute("tabindex", "0");
    head.setAttribute("aria-expanded", "true");

    const toggle = () => {
      const collapsed = section.classList.toggle("is-collapsed");
      head.setAttribute("aria-expanded", collapsed ? "false" : "true");
    };
    head.addEventListener("click", (event) => {
      if (event.target.closest("a")) return;
      toggle();
    });
    head.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        toggle();
      }
    });
  }
}

function expandScenarioFlowFromHash() {
  const id = window.location.hash.slice(1);
  if (!id) return;
  const target = document.getElementById(id);
  if (!target) return;
  const section = target.closest("section.scenario-flow.is-collapsed");
  if (!section) return;
  section.classList.remove("is-collapsed");
  const head = section.querySelector(":scope > .section-head");
  if (head) head.setAttribute("aria-expanded", "true");
}

function initScenarioEnhancements() {
  makeScenarioFlowsCollapsible();
  buildScenarioToc();
  expandScenarioFlowFromHash();
}

window.addEventListener("hashchange", expandScenarioFlowFromHash);

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initScenarioEnhancements, { once: true });
} else {
  initScenarioEnhancements();
}