import * as fs from "fs";
import * as path from "path";
import * as process from "process";

const hexToCssVarMap: { [key: string]: string } = {
  white: "var(--vscode-editor-background, white)",
  black: "var(--vscode-editor-foreground, black)",
  "#505050": "var(--vscode-editor-foreground, #505050)",
  "#242424": "var(--vscode-editor-foreground, #242424)",
  "#EEECEC": "var(--vscode-editorGroupHeader-tabsBackground, #EEECEC)",
  "#D9D9D9": "var(--vscode-editorGroupHeader-tabsBackground, #D9D9D9)",
  "#FFFFFF": "var(--vscode-editor-foreground, #FFFFFF)",
  "#F8F8F8": "var(--vscode-editorGroupHeader-tabsBackground, #F8F8F8)",
  "#E5E5E5": "var(--vscode-activityBar-border, #E5E5E5)",
  "#616161": "var(--vscode-badge-background, #616161)",
  "#005FB8": "var(--vscode-panelTitle-activeBorder, #005FB8)",
  "#868686": "var(--vscode-input-placeholderForeground, #868686)",
  "#CCCCCC": "var(--vscode-badge-background, #CCCCCC)",
  "#D2ECFF": "var(--vscode-chat-slashCommandBackground, #D2ECFF)",
  "#3B3B3B": "var(--vscode-icon-foreground, #3B3B3B)",
  "#3C3C3C": "var(--vscode-titleBar-activeBackground, #3C3C3C)",
  "#333333": "var(--vscode-titleBar-activeBackground, #333333)",
  "#E7E7E7": "var(--vscode-activityBar-border, #E7E7E7)",
  "#ADD6FF": "var(--vscode-badge-background, #ADD6FF)",
  "#DDDDDD": "var(--vscode-actionBar-toggledBackground, #DDDDDD)",
  "#F85149": "var(--vscode-errorForeground, #F85149)",
  "#2c2c2d": "var(--vscode-notificationCenterHeader-background, #2c2c2d)",
  "#252526": "var(--vscode-editorWidget-background, #252526)",
  "#292929": "var(--vscode-editor-background, #292929)",
  "#007ACC": "var(--vscode-button-background, #007ACC)",
  "#1E1E1E": "var(--vscode-editor-background, #1E1E1E)",
  "#007FD4": "var(--vscode-button-background, #007FD4)",
  "#9D9D9D": "var(--vscode-input-placeholderForeground, #9D9D9D)",
  "#062F4A": "var(--vscode-list-activeSelectionBackground, #062F4A)",
  "#E3E3E3": "var(--vscode-checkbox-foreground, #E3E3E3)",
  "#979797": "var(--vscode-activityBar-border, #979797)",
  "#C4C4C4": "var(--vscode-activityBar-border, #C4C4C4)",
  "#0078D4": "var(--vscode-button-background, #0078D4)",
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
