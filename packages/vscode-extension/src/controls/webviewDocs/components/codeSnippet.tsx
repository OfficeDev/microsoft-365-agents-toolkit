import "./codeSnippet.scss";

import * as React from "react";
import { CopyToClipboard } from "react-copy-to-clipboard";
import SyntaxHighlighter from "react-syntax-highlighter";
import { lightfair, dark } from "react-syntax-highlighter/dist/esm/styles/hljs";

import { CheckmarkRegular, CopyRegular } from "@fluentui/react-icons";

import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../../../telemetry/extTelemetryEvents";
import { Commands } from "../../Commands";

export default function CodeSnippet(props: {
  data: string;
  language: string;
  theme: string;
  identifier: string;
}) {
  const [copied, setCopied] = React.useState(false);

  React.useEffect(() => {
    const timeout = setTimeout(() => {
      if (copied) {
        setCopied(false);
      }
    }, 1000);

    return () => clearTimeout(timeout);
  }, [copied]);

  const onCopyCode = () => {
    vscode.postMessage({
      command: Commands.SendTelemetryEvent,
      data: {
        eventName: TelemetryEvent.CopyCodeSnippet,
        properties: {
          [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.InProductDoc,
          [TelemetryProperty.Identifier]: props.identifier,
        },
      },
    });
    setCopied(true);
  };

  let copyButton;
  if (copied) {
    copyButton = (
      <div className="copiedButton">
        <CheckmarkRegular />
      </div>
    );
  } else {
    copyButton = (
      <div className="copyButton">
        <CopyRegular />
        <button>Copy</button>
      </div>
    );
  }

  return (
    <div className="codeSnippet">
      <div className="codeTitle">
        <CopyToClipboard text={props.data} onCopy={onCopyCode}>
          {copyButton}
        </CopyToClipboard>
      </div>
      <div>
        <SyntaxHighlighter
          language={props.language}
          style={props.theme === "light" ? lightfair : dark}
          className="codeBlock"
        >
          {props.data}
        </SyntaxHighlighter>
      </div>
    </div>
  );
}
