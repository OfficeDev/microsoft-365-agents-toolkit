import * as fs from "fs";
import * as path from "path";
import * as process from "process";

const hexToCssVarMap: { [key: string]: string } = {
  white: "var(--vscode-editor-background, white)",
  "#FFFFFF": "var(--vscode-editor-background, white)",
  "#F8F8F8": "var(--vscode-editorGroupHeader-tabsBackground, #F8F8F8)",
  "#E5E5E5": "var(--vscode-activityBar-border, #E5E5E5)",
  "#616161": "var(--vscode-badge-background, #616161)",
  "#005FB8": "var(--vscode-panelTitle-activeBorder, #005FB8)",
  "#868686": "var(--vscode-input-placeholderForeground, #868686)",
  "#CCCCCC": "var(--vscode-menu-foreground, #CCCCCC)",
  "#D2ECFF": "var(--vscode-chat-slashCommandBackground, #D2ECFF)",
  "#3B3B3B": "var(--vscode-icon-foreground, #3B3B3B)",
  "#E7E7E7": "var(--vscode-editorGroupHeader-tabsBorder, #E7E7E7)",
  "#ADD6FF": "var(--vscode-editor-selectionHighlightBackground, #ADD6FF)",
  "#DDDDDD": "var(--vscode-actionBar-toggledBackground, #DDDDDD)",
  "#F85149": "var(--vscode-errorForeground, #F85149)",
  "#2c2c2d": "var(--vscode-notificationCenterHeader-background, #2c2c2d)",
  "#252526": "var(--vscode-editorWidget-background, #252526)",
};

function replaceHexWithCssVar(content: string): string {
  for (const [hex, cssVar] of Object.entries(hexToCssVarMap)) {
    const regex = new RegExp(hex, "gi");
    content = content.replace(regex, cssVar);
  }
  return content;
}

function processSvgFile(filePath: string): void {
  if (path.extname(filePath) === ".svg") {
    fs.readFile(filePath, "utf8", (err, data) => {
      if (err) {
        console.error("Error reading file:", err);
        return;
      }

      const updatedContent = replaceHexWithCssVar(data);
      fs.writeFile(filePath, updatedContent, "utf8", (err) => {
        if (err) {
          console.error("Error writing file:", err);
        } else {
          console.log(`Processed ${filePath}`);
        }
      });
    });
  } else {
    console.error("The provided file is not an SVG file.");
  }
}

// Get file path from command line input
const filePath = process.argv[2];
if (filePath) {
  processSvgFile(filePath);
} else {
  console.error("Please provide a file path as a command line argument.");
}
