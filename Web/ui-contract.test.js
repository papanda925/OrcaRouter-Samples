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
  !sendHandler.includes('renderConversation(question, "");'),
  "Send must not flash an unconfirmed Question in Conversation."
);

assert(
  sendHandler.includes("renderConversation();"),
  "Send must keep committed conversation visible while waiting."
);

assert(
  sendHandler.includes("renderConversation(question, errorText);"),
  "Errors must keep the submitted Question and error visible."
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
  /if \(pendingAssistant\) \{\s*messages\.push\(\{ role: "assistant", content: pendingAssistant \}\);/m.test(app),
  "Validation errors must remain visible even when Question is empty."
);

console.log("Web UI behavior contract checks passed.");
