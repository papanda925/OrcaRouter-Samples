const fs = require("fs");
const assert = require("assert");

const html = fs.readFileSync("Web/index.html", "utf8");
const app = fs.readFileSync("Web/app.js", "utf8");

function extractHandler(source, startMarker, endMarker) {
  const start = source.indexOf(startMarker);
  assert.notStrictEqual(start, -1, `Missing handler: ${startMarker}`);

  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.notStrictEqual(end, -1, `Missing handler end after: ${startMarker}`);

  return source.slice(start, end);
}

assert(
  html.includes('id="applyPromptExampleButton"'),
  "Web UI must have an explicit prompt-example apply button."
);

assert(
  html.includes(
    'href="https://github.com/papanda925/OrcaRouter-Samples/blob/main/docs/processing-flow.md"'
  ),
  "Common flow must use a URL that works from the Web-only local server."
);

assert(
  !html.includes('href="../docs/processing-flow.md"'),
  "Common flow must not point outside the Web server root."
);

const promptSelection = extractHandler(
  app,
  'promptExampleInput.addEventListener("change"',
  'applyPromptExampleButton.addEventListener("click"'
);

assert(
  !promptSelection.includes("questionInput.value ="),
  "Selecting a prompt example must not overwrite Question."
);

const promptApply = extractHandler(
  app,
  'applyPromptExampleButton.addEventListener("click"',
  'newChatButton.addEventListener("click"'
);

assert(
  promptApply.includes("questionInput.value = prompt"),
  "Prompt example must be copied only by the explicit apply action."
);

const modeHandler = extractHandler(
  app,
  'modeInput.addEventListener("change"',
  'apiKeyInput.addEventListener("input"'
);

assert(
  !modeHandler.includes("questionInput.value ="),
  "Changing Mode must not overwrite Question."
);

const sendHandler = extractHandler(
  app,
  "async function sendChat()",
  'modeInput.addEventListener("change"'
);

assert(
  sendHandler.includes('renderAnswer("");'),
  "Send must clear the previous Answer while waiting for the new response."
);

assert(
  !sendHandler.includes("renderConversation"),
  "The Web result area must not render the conversation transcript."
);

assert(
  sendHandler.includes("renderAnswer(errorText);"),
  "Errors must remain visible in the Answer area without repeating Question."
);

assert(
  sendHandler.includes("Developer / Trace"),
  "Error status must direct users to optional Developer / Trace diagnostics."
);

assert(
  app.includes("responseBody: parsed ?? rawBody ??"),
  "HTTP errors must preserve response data for Developer Information."
);

assert(
  app.includes("function renderAnswer(answerText"),
  "Web must have a dedicated answer-only renderer."
);

assert(
  app.includes("setUiBusy(true)") &&
    app.includes("setUiBusy(false)") &&
    html.includes('id="busyIndicator"'),
  "Web must show and clear a busy state around requests."
);

assert(
  !app.includes("conversation-message"),
  "Web result rendering must not recreate prior user/assistant transcript blocks."
);

console.log("Web UI behavior contract checks passed.");
