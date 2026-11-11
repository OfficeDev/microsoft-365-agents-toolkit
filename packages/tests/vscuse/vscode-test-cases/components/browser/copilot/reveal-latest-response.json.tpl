{
  "component": {
    "version": 1,
    "uiSurface": "browser",
    "hostSurface": "copilot",
    "id": "revealLatestCopilotResponse",
    "parameters": ["instanceSuffix"],
    "recording": {
      "report": "da-mcp-dcr-update-copilot",
      "stepId": "step_assertChatNonEmptyField_caceeb9ea_13_10",
      "screenshot": "screenshot_20261111_100130_644.jpg",
      "preconditions": [
        "dhash:627:595:16:5:2d0e6d29874e2d92",
        "dhash:627:595:96:5:0008140d0c389906"
      ]
    }
  },
  "steps": [
    {
      "step_id": "step_revealLatestCopilotResponse_click_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": { "button": "left", "x": 627, "y": 595 },
      "description": "Click the conversation's circular down-arrow Scroll to bottom button above the message composer in Microsoft 365 Copilot to reveal the latest completed assistant response, not the browser scrollbar, sidebar, or an earlier Allow card.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:browser",
        "host_surface:copilot",
        "entry_state:completed-reply",
        "action:reveal-latest-response",
        "ocr:true",
        "source_screenshot:screenshot_20261111_100130_644.jpg"
      ]
    }
  ]
}
