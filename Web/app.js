const API_ENDPOINT = "https://api.orcarouter.ai/v1/chat/completions";
const API_KEY_PLACEHOLDER = "xxx-your-orcarouter-api-key-xxx";
const REQUEST_TIMEOUT_MS = 60000;

const apiKeyInput = document.getElementById("apiKey");
const modelInput = document.getElementById("model");
const questionInput = document.getElementById("question");
const answerBox = document.getElementById("answer");
const traceList = document.getElementById("trace");
const sendButton = document.getElementById("sendButton");
const clearTraceButton = document.getElementById("clearTraceButton");
const statusText = document.getElementById("statusText");
const traceItemTemplate = document.getElementById("traceItemTemplate");

function setStatus(message, isError = false) {
  statusText.textContent = message;
  statusText.classList.toggle("error", isError);
}

function maskApiKey(apiKey) {
  if (!apiKey) {
    return "(empty)";
  }

  if (apiKey.length <= 8) {
    return "********";
  }

  return `${apiKey.slice(0, 4)}...${apiKey.slice(-4)}`;
}

function formatTraceData(data) {
  if (typeof data === "string") {
    return data;
  }

  try {
    return JSON.stringify(data, null, 2);
  } catch {
    return String(data);
  }
}

function addTrace(step, direction, title, data = "") {
  const fragment = traceItemTemplate.content.cloneNode(true);
  const item = fragment.querySelector(".trace-item");

  fragment.querySelector(".trace-step").textContent = step;
  fragment.querySelector(".trace-direction").textContent = direction;
  fragment.querySelector(".trace-time").textContent =
    new Date().toLocaleTimeString("ja-JP", { hour12: false });
  fragment.querySelector(".trace-title").textContent = title;
  fragment.querySelector(".trace-data").textContent = formatTraceData(data);

  traceList.appendChild(fragment);
  item?.scrollIntoView({ block: "nearest", behavior: "smooth" });
}

function clearTrace() {
  traceList.innerHTML = "";
  const empty = document.createElement("p");
  empty.className = "trace-empty";
  empty.textContent = "送信すると、6つの共通処理ステップとHTTP電文がここに表示されます。";
  traceList.appendChild(empty);
}

function clearTracePlaceholder() {
  traceList.querySelector(".trace-empty")?.remove();
}

function extractAssistantText(json) {
  const content = json?.choices?.[0]?.message?.content;

  if (typeof content === "string") {
    return content;
  }

  if (Array.isArray(content)) {
    const textParts = content
      .filter((part) => part && part.type === "text" && typeof part.text === "string")
      .map((part) => part.text);

    if (textParts.length > 0) {
      return textParts.join("\n");
    }
  }

  throw new Error(
    "choices[0].message.content が見つかりません。Trace の Raw response を確認してください。"
  );
}

async function sendChat() {
  const apiKey = apiKeyInput.value.trim();
  const model = modelInput.value.trim();
  const question = questionInput.value.trim();

  clearTracePlaceholder();
  answerBox.textContent = "";
  setStatus("Processing...");
  sendButton.disabled = true;

  const startedAt = performance.now();
  let rawResponse = "";
  let responseStatus = null;

  try {
    // STEP 1: Validate inputs.
    addTrace("STEP 1", "LOCAL", "入力値を検証", {
      apiKey: maskApiKey(apiKey),
      model,
      questionLength: question.length
    });

    if (!apiKey || apiKey === API_KEY_PLACEHOLDER || apiKey.startsWith("xxx-")) {
      throw new Error(
        "APIキーがダミー値のままです。OrcaRouterで発行したAPIキーを入力してください。"
      );
    }

    if (!model) {
      throw new Error("Model を入力してください。");
    }

    if (!question) {
      throw new Error("質問を入力してください。");
    }

    // STEP 2: Build request.
    const requestBody = {
      model,
      messages: [
        {
          role: "user",
          content: question
        }
      ]
    };

    addTrace("STEP 2", "REQUEST", "HTTPリクエストを組み立て", {
      method: "POST",
      endpoint: API_ENDPOINT,
      headers: {
        Authorization: `Bearer ${maskApiKey(apiKey)}`,
        "Content-Type": "application/json"
      },
      body: requestBody
    });

    // STEP 3: Send HTTP POST.
    const controller = new AbortController();
    const timeoutId = window.setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

    addTrace("STEP 3", "REQUEST", "OrcaRouterへPOSTを送信", {
      timeoutMs: REQUEST_TIMEOUT_MS
    });

    let response;

    try {
      response = await fetch(API_ENDPOINT, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify(requestBody),
        signal: controller.signal
      });
    } finally {
      window.clearTimeout(timeoutId);
    }

    // STEP 4: Receive response.
    responseStatus = response.status;
    rawResponse = await response.text();
    const elapsedMs = Math.round(performance.now() - startedAt);
    const responseHeaders = Object.fromEntries(response.headers.entries());

    addTrace("STEP 4", "RESPONSE", "HTTPレスポンスを受信", {
      status: response.status,
      statusText: response.statusText,
      elapsedMs,
      headers: responseHeaders,
      rawBody: rawResponse
    });

    if (!response.ok) {
      throw new Error(
        `HTTP ${response.status} ${response.statusText}: ${rawResponse || "(empty response)"}`
      );
    }

    // STEP 5: Parse assistant message.
    let json;

    try {
      json = JSON.parse(rawResponse);
    } catch (error) {
      throw new Error(`レスポンスJSONの解析に失敗しました: ${error.message}`);
    }

    const assistantText = extractAssistantText(json);

    addTrace("STEP 5", "LOCAL", "Assistantメッセージを解析", {
      answerLength: assistantText.length,
      usage: json.usage ?? "(usage not returned)"
    });

    // STEP 6: Update UI and trace.
    answerBox.textContent = assistantText;

    addTrace("STEP 6", "LOCAL", "画面へ回答を表示", {
      completed: true,
      totalElapsedMs: Math.round(performance.now() - startedAt)
    });

    setStatus("Completed");
  } catch (error) {
    const isAbort = error?.name === "AbortError";
    const message = isAbort
      ? `タイムアウトしました（${REQUEST_TIMEOUT_MS / 1000}秒）。`
      : error?.message || String(error);

    addTrace("ERROR", "ERROR", "処理中にエラーが発生", {
      message,
      name: error?.name ?? "(unknown)",
      stack: error?.stack ?? "(stack not available)",
      httpStatus: responseStatus ?? "(no HTTP response)",
      rawResponse: rawResponse || "(not available)"
    });

    answerBox.textContent = `ERROR: ${message}`;
    setStatus("Error - Trace を確認してください", true);
  } finally {
    sendButton.disabled = false;
  }
}

sendButton.addEventListener("click", sendChat);

clearTraceButton.addEventListener("click", () => {
  clearTrace();
  setStatus("Ready");
});

questionInput.addEventListener("keydown", (event) => {
  if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
    sendChat();
  }
});

clearTrace();
