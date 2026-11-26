import json
import sys
import time
from urllib.parse import urlsplit
from urllib.request import urlopen

from websockets.sync.client import connect


# The DA's "Launch Agent in Copilot (Chrome)" profile exposes this loopback port.
DEBUG_PORT = 9223
COPILOT_HOSTS = {"copilot.cloud.microsoft", "m365.cloud.microsoft"}

CHECK = """(agentName => {
  if (document.visibilityState !== 'visible') return 'waiting';
  const visible = el => el.checkVisibility({opacityProperty: true, visibilityProperty: true});
  const candidates = [...document.querySelectorAll('[role="dialog"] button')].filter(button => {
    if (button.getAttribute('aria-label') !== 'Add ' + agentName ||
        button.innerText.trim() !== 'Add' || !visible(button) ||
        button.matches(':disabled, [aria-disabled="true"]')) return false;
    const rect = button.getBoundingClientRect();
    return button.contains(document.elementFromPoint(rect.x + rect.width / 2, rect.y + rect.height / 2));
  });
  if (candidates.length === 1) {
    candidates[0].click();
    return 'clicked';
  }
  if (![...document.querySelectorAll('[role="dialog"]')].some(visible) &&
      [...document.querySelectorAll('textarea, [contenteditable="true"]')].some(visible)) {
    return 'chat';
  }
  return 'waiting';
})"""


def try_add(agent_name):
    deadline = time.monotonic() + 15
    while time.monotonic() < deadline:
        with urlopen(f"http://127.0.0.1:{DEBUG_PORT}/json/list", timeout=3) as response:
            pages = [
                page for page in json.load(response)
                if page["type"] == "page" and urlsplit(page["url"]).scheme == "https"
                and urlsplit(page["url"]).hostname in COPILOT_HOSTS
            ]
        if len(pages) == 1:
            page = pages[0]
            with connect(f"ws://127.0.0.1:{DEBUG_PORT}/devtools/page/{page['id']}", open_timeout=3) as ws:
                expression = (
                    "['https://copilot.cloud.microsoft','https://m365.cloud.microsoft'].includes(location.origin)"
                    f" ? {CHECK}({json.dumps(agent_name)}) : 'waiting'"
                )
                ws.send(json.dumps({"id": 1, "method": "Runtime.evaluate",
                                    "params": {"expression": expression, "returnByValue": True}}))
                reply = json.loads(ws.recv(timeout=3))
                if "error" in reply or "exceptionDetails" in reply["result"]:
                    raise RuntimeError("Could not inspect the local Copilot preview")
                outcome = reply["result"]["result"]["value"]
                if outcome != "waiting":
                    return outcome
        time.sleep(0.5)
    return "not offered"


if __name__ == "__main__":
    print("Copilot optional Add:", try_add(sys.argv[1]))
