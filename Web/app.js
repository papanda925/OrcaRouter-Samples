const API_ENDPOINT = "https://api.orcarouter.ai/v1/chat/completions";
const API_KEY_PLACEHOLDER = "xxx-your-orcarouter-api-key-xxx";

// LOCAL TEST ONLY:
// 実APIキーをソースへ一時的に埋め込んで試す場合は、次の値だけを書き換えます。
// 公開GitHubへ実APIキーをコミットしないでください。
const DEFAULT_API_KEY = "xxx-your-orcarouter-api-key-xxx";

const CHAT_TOOL_TIMEOUT_MS = 120000;
const STREAM_TIMEOUT_MS = 90000;
const MAX_STREAM_TRACE_EVENTS = 50;
const MAX_RAW_JSON_CHARS = 30000;
const MAX_HISTORY_TURNS = 10;

const apiKeyInput = document.getElementById("apiKey");
const apiKeyFileInput = document.getElementById("apiKeyFile");
const modelInput = document.getElementById("model");
const modeInput = document.getElementById("mode");
const questionInput = document.getElementById("question");
const answerBox = document.getElementById("answer");
const rawJsonTitle = document.getElementById("rawJsonTitle");
const rawJsonStatus = document.getElementById("rawJsonStatus");
const rawJsonBox = document.getElementById("rawJson");
const traceList = document.getElementById("trace");
const sendButton = document.getElementById("sendButton");
const clearTraceButton = document.getElementById("clearTraceButton");
const statusText = document.getElementById("statusText");
const traceItemTemplate = document.getElementById("traceItemTemplate");
const newChatButton = document.getElementById("newChatButton");
const promptExampleInput = document.getElementById("promptExample");
const applyPromptExampleButton = document.getElementById("applyPromptExampleButton");
const historyStatus = document.getElementById("historyStatus");
const firstRunHelp = document.getElementById("firstRunHelp");
const devHttpStatus = document.getElementById("devHttpStatus");
const devElapsed = document.getElementById("devElapsed");
const devModel = document.getElementById("devModel");
const devPromptTokens = document.getElementById("devPromptTokens");
const devCompletionTokens = document.getElementById("devCompletionTokens");
const devTotalTokens = document.getElementById("devTotalTokens");
const devCost = document.getElementById("devCost");
const devHistory = document.getElementById("devHistory");
const devRequestJson = document.getElementById("devRequestJson");
const devResponseJson = document.getElementById("devResponseJson");

const conversationHistory = [];
let transientAssistantText = "";

apiKeyInput.value = DEFAULT_API_KEY;

const PROMPT_EXAMPLES = {
  summary: "次の文章を3行で要約してください。\n\nここに文章を貼り付けてください。",
  explain: "次の内容を、専門用語を補足しながら初心者向けに説明してください。\n\nここに内容を貼り付けてください。",
  review: "次のコードをレビューし、問題点・理由・改善例の順に説明してください。\n\nここにコードを貼り付けてください。",
  json: "次の内容を整理し、JSON形式だけで返してください。\n\nここに内容を貼り付けてください。",
  translate: "次の日本語を自然な英語に翻訳してください。\n\nここに文章を貼り付けてください。"
};

function hasConfiguredApiKey() {
  const value = apiKeyInput.value.trim();
  return Boolean(value) &&
    value !== API_KEY_PLACEHOLDER &&
    !value.startsWith("xxx-");
}

function updateFirstRunHelp() {
  firstRunHelp.hidden = hasConfiguredApiKey();
}

function getConversationMessages(question) {
  const messages = [];

  for (const turn of conversationHistory) {
    messages.push({ role: "user", content: turn.user });
    messages.push({ role: "assistant", content: turn.assistant });
  }

  messages.push({ role: "user", content: question });
  return messages;
}

function trimConversationHistory() {
  while (conversationHistory.length > MAX_HISTORY_TURNS) {
    conversationHistory.shift();
  }
}

function updateHistoryStatus() {
  const count = conversationHistory.length;
  historyStatus.textContent = `履歴 ${count} / ${MAX_HISTORY_TURNS} 往復`;
  devHistory.textContent = `${count} / ${MAX_HISTORY_TURNS} turns`;
}

function renderConversation(pendingQuestion = "", pendingAssistant = "") {
  answerBox.innerHTML = "";

  const messages = [];
  for (const turn of conversationHistory) {
    messages.push({ role: "user", content: turn.user });
    messages.push({ role: "assistant", content: turn.assistant });
  }

  if (pendingQuestion) {
    messages.push({ role: "user", content: pendingQuestion });
    if (pendingAssistant) {
      messages.push({ role: "assistant", content: pendingAssistant });
    }
  }

  if (messages.length === 0) {
    answerBox.textContent = "ここに会話が表示されます。";
    return;
  }

  for (const message of messages) {
    const item = document.createElement("div");
    item.className = `conversation-message ${message.role}`;

    const role = document.createElement("span");
    role.className = "conversation-role";
    role.textContent = message.role === "user" ? "YOU" : "ASSISTANT";

    const body = document.createElement("div");
    body.textContent = message.content;

    item.append(role, body);
    answerBox.appendChild(item);
  }

  answerBox.scrollTop = answerBox.scrollHeight;
}

function addConversationTurn(question, assistant) {
  conversationHistory.push({ user: question, assistant });
  trimConversationHistory();
  updateHistoryStatus();
  renderConversation();
}

function startNewChat() {
  conversationHistory.length = 0;
  transientAssistantText = "";
  updateHistoryStatus();
  renderConversation();
  setStatus("New chat - 履歴をクリアしました");
}

function combineUsage(...items) {
  const result = {
    prompt_tokens: 0,
    completion_tokens: 0,
    total_tokens: 0,
    cost_usd: 0
  };
  let hasAny = false;
  let hasCost = false;

  for (const usage of items) {
    if (!usage || typeof usage !== "object") continue;
    hasAny = true;

    for (const key of ["prompt_tokens", "completion_tokens", "total_tokens"]) {
      const value = Number(usage[key]);
      if (Number.isFinite(value)) result[key] += value;
    }

    const cost = Number(usage.cost_usd);
    if (Number.isFinite(cost)) {
      result.cost_usd += cost;
      hasCost = true;
    }
  }

  if (!hasAny) return null;
  if (!hasCost) delete result.cost_usd;
  return result;
}

function updateDeveloperRequest(value) {
  devRequestJson.textContent = JSON.stringify(value ?? {}, null, 2);
}

function updateDeveloperResponse(value) {
  if (typeof value === "string") {
    devResponseJson.textContent = limitRawText(value);
    return;
  }
  devResponseJson.textContent = JSON.stringify(value ?? {}, null, 2);
}

function updateDeveloperInfo({ status, elapsedMs, model, usage, request, response }) {
  devHttpStatus.textContent = status ? String(status) : "-";
  devElapsed.textContent = Number.isFinite(elapsedMs) ? `${Math.round(elapsedMs)} ms` : "-";
  devModel.textContent = model || "-";
  devPromptTokens.textContent = usage?.prompt_tokens ?? "-";
  devCompletionTokens.textContent = usage?.completion_tokens ?? "-";
  devTotalTokens.textContent = usage?.total_tokens ?? "-";
  devCost.textContent =
    typeof usage?.cost_usd === "number"
      ? `${usage.cost_usd.toFixed(6)}`
      : "(not returned)";
  updateDeveloperRequest(request);
  updateDeveloperResponse(response);
  updateHistoryStatus();
}

function resetDeveloperInfo() {
  updateDeveloperInfo({
    status: null,
    elapsedMs: NaN,
    model: modelInput.value.trim(),
    usage: null,
    request: {},
    response: {}
  });
}

function setStatus(message, isError = false) {
  statusText.textContent = message;
  statusText.classList.toggle("error", isError);
}

function maskApiKey(apiKey) {
  if (!apiKey) return "(empty)";
  if (apiKey.length <= 8) return "********";
  return `${apiKey.slice(0, 4)}...${apiKey.slice(-4)}`;
}

function extractApiKeyFromText(text) {
  const match = text.match(/sk-orca-[A-Za-z0-9._-]+/);

  if (!match) {
    throw new Error(
      "選択したファイルから完全な OrcaRouter API Key（sk-orca-...）を見つけられませんでした。"
    );
  }

  return match[0];
}

async function loadApiKeyFromFile(file) {
  if (!file) return;

  const maxBytes = 1024 * 1024;

  if (file.size > maxBytes) {
    throw new Error("キーのファイルが大きすぎます。1MB以下のファイルを選択してください。");
  }

  const text = await file.text();
  const apiKey = extractApiKeyFromText(text);

  apiKeyInput.value = apiKey;

  clearTracePlaceholder();
  addTrace("KEY", "LOCAL", "APIキーをローカルファイルから読み込み", {
    source: "(local file selected)",
    fileSize: file.size,
    apiKey: maskApiKey(apiKey),
    note: "ファイル名・ファイル本文・完全なAPIキーはTraceへ出力していません。"
  });

  setStatus("API Key loaded from local file");
  updateFirstRunHelp();
}

function limitRawText(value) {
  const text = String(value ?? "");
  if (text.length <= MAX_RAW_JSON_CHARS) return text;
  return `${text.slice(0, MAX_RAW_JSON_CHARS)}\n...(truncated for readability)`;
}

function prepareRawResponse(caption = "Raw JSON") {
  rawJsonTitle.textContent = caption;
  rawJsonStatus.textContent = "Waiting for HTTP response...";
  rawJsonBox.textContent = "";
}

function displayRawResponse(rawText, caption, status = 0) {
  rawJsonTitle.textContent = caption;
  rawJsonStatus.textContent =
    `${status ? `HTTP Status: ${status}` : "HTTP Status: (not available)"} / Chars: ${String(rawText ?? "").length}`;
  rawJsonBox.textContent = limitRawText(rawText);
}

function formatTraceData(data) {
  if (typeof data === "string") return data;
  try {
    return JSON.stringify(data, null, 2);
  } catch {
    return String(data);
  }
}

function addTrace(step, direction, title, data = "") {
  const fragment = traceItemTemplate.content.cloneNode(true);
  fragment.querySelector(".trace-step").textContent = step;
  fragment.querySelector(".trace-direction").textContent = direction;
  fragment.querySelector(".trace-time").textContent =
    new Date().toLocaleTimeString("ja-JP", { hour12: false });
  fragment.querySelector(".trace-title").textContent = title;
  fragment.querySelector(".trace-data").textContent = formatTraceData(data);
  traceList.appendChild(fragment);
}

function clearTrace() {
  traceList.innerHTML = "";
  const empty = document.createElement("p");
  empty.className = "trace-empty";
  empty.textContent =
    "送信すると、共通6ステップに沿って Request / Response / SSE / Tool Calling の流れを表示します。";
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

function buildErrorDetails(status, statusText, headers, rawBody) {
  let parsed = null;

  try {
    parsed = rawBody ? JSON.parse(rawBody) : null;
  } catch {
    parsed = null;
  }

  const error = parsed?.error ?? {};
  const retryAfter = headers?.["retry-after"] ?? null;
  const code = error.code ?? "";
  const type = error.type ?? "";
  const message = error.message ?? rawBody ?? statusText ?? "Unknown error";

  let guidance = "Trace の error.code / error.type / HTTP Status を確認してください。";

  if (code === "free_quota_exhausted") {
    guidance =
      "orcarouter/free の無料枠または利用可能な無料モデルがありません。Tool Calling等の処理へ到達する前にAPI側で拒否されています。有料モデルへは自動切替しません。";
  } else if (status === 400 && code === "bad_request_body") {
    guidance = "Request JSONを解析できません。Traceのrequest bodyを確認してください。";
  } else if (status === 400 && code === "model_price_error") {
    guidance = "選択モデルの価格設定に問題があります。OrcaRouter側のモデル情報を確認してください。";
  } else if (status === 400 && code === "api_not_implemented") {
    guidance = "選択モデルではこのAPI操作がサポートされていません。";
  } else if (
    status === 400 &&
    ["prompt_blocked", "sensitive_words_detected", "guardrail_blocked"].includes(code)
  ) {
    guidance = "Providerの安全ポリシーまたはWorkspace guardrailで拒否されています。";
  } else if (status === 400 && code === "firewall_blocked") {
    guidance = "Agent FirewallがToolを拒否しました。Firewall policyとmetadataを確認してください。";
  } else if (status === 400 && code === "firewall_approval_pending") {
    guidance = "Tool CallがFirewall承認待ちです。単純な再試行では解消しません。";
  } else if (status === 401) {
    guidance = "APIキーが無効、またはAuthorizationヘッダーが不正です。";
  } else if (status === 402) {
    guidance = "支払いまたはQuotaが必要です。error.code と残高/無料枠を確認してください。";
  } else if (status === 403 && code === "insufficient_user_quota") {
    guidance = "Workspace残高またはメンバー/エージェント予算を確認してください。";
  } else if (status === 403 && code === "pre_consume_token_quota_failed") {
    guidance = "APIキー自身のQuota上限を確認してください。";
  } else if (status === 403 && code === "access_denied") {
    guidance = "APIキーは認識されていますが、このリクエストは許可されていません。モデル権限・IP制限・利用上限を確認してください。";
  } else if (status === 404) {
    guidance = "EndpointまたはModel IDが見つかりません。指定値を確認してください。";
  } else if (status === 425) {
    guidance = "指定モデルはまだ利用開始前の可能性があります。error.metadata も確認してください。";
  } else if (status === 429 && retryAfter) {
    guidance = `Rate Limitです。Retry-After=${retryAfter} 秒を待ってから再試行してください。`;
  } else if (status === 429) {
    guidance = "Retry-Afterがない無料枠429は、同じ長いPromptを待って再送しても改善しない場合があります。";
  } else if (status === 500) {
    guidance = "OrcaRouter内部エラーです。時間を置いて再試行してください。";
  } else if (status === 502) {
    guidance = "上流Providerまたはfallback routeが失敗しています。時間を置くかTraceのHeaderを確認してください。";
  } else if (status === 503 && code === "model_not_found") {
    guidance = "そのモデルが現在のアカウントで利用可能か確認してください。";
  } else if (status === 503 && code === "byok:key_unavailable") {
    guidance = "WorkspaceのBYOK provider keyを利用できません。Provider keyまたはfallback設定を確認してください。";
  } else if (status === 503) {
    guidance = "OrcaRouterまたは上流Providerが一時的に利用できない可能性があります。";
  }

  return {
    httpStatus: status,
    httpStatusText: statusText,
    errorType: type,
    errorCode: code,
    errorMessage: message,
    retryAfter,
    metadata: error.metadata ?? null,
    guidance,
    headers: headers ?? {},
    responseBody: parsed ?? rawBody ?? ""
  };
}

function validateInputs(apiKey, model, question, mode) {
  addTrace("STEP 1", "LOCAL", "入力値を検証", {
    apiKey: maskApiKey(apiKey),
    model,
    mode,
    questionLength: question.length
  });

  if (!apiKey || apiKey === API_KEY_PLACEHOLDER || apiKey.startsWith("xxx-")) {
    throw new Error(
      "APIキーがダミー値のままです。OrcaRouterで発行したAPIキーを入力してください。"
    );
  }

  if (!model) throw new Error("Model を入力してください。");
  if (!question) throw new Error("質問を入力してください。");
}

function createAbortController(timeoutMs) {
  const controller = new AbortController();
  const timeoutId = window.setTimeout(() => controller.abort(), timeoutMs);
  return { controller, timeoutId };
}

async function postJson(
  apiKey,
  body,
  startedAt,
  traceTitle = "OrcaRouterへPOSTを送信",
  rawCaption = "Raw JSON",
  timeoutMs = CHAT_TOOL_TIMEOUT_MS
) {
  addTrace("STEP 3", "REQUEST", traceTitle, {
    timeoutMs
  });

  const { controller, timeoutId } = createAbortController(timeoutMs);

  try {
    const response = await fetch(API_ENDPOINT, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json; charset=utf-8",
        Accept: "application/json",
        "X-OrcaRouter-Include-Cost": "true"
      },
      body: JSON.stringify(body),
      signal: controller.signal
    });

    const rawBody = await response.text();
    const headers = Object.fromEntries(response.headers.entries());

    displayRawResponse(rawBody, rawCaption, response.status);

    addTrace("STEP 4", "RESPONSE", "HTTPレスポンスを受信", {
      status: response.status,
      statusText: response.statusText,
      elapsedMs: Math.round(performance.now() - startedAt),
      headers,
      rawBody
    });

    if (!response.ok) {
      const details = buildErrorDetails(
        response.status,
        response.statusText,
        headers,
        rawBody
      );
      const error = new Error(
        `HTTP ${response.status}: ${details.errorMessage}\n${details.guidance}`
      );
      error.details = details;
      throw error;
    }

    let json;

    try {
      json = rawBody ? JSON.parse(rawBody) : {};
    } catch (parseError) {
      const error = new Error(
        `HTTP ${response.status} のレスポンスJSONを解析できません: ${parseError.message}`
      );
      error.details = {
        httpStatus: response.status,
        httpStatusText: response.statusText,
        headers,
        responseBody: rawBody,
        guidance: "Raw JSON / Trace を確認してください。"
      };
      throw error;
    }

    return {
      response,
      headers,
      rawBody,
      json
    };
  } finally {
    window.clearTimeout(timeoutId);
  }
}

async function runChat(apiKey, model, question, startedAt) {
  // STEP 2: Build request. The application owns the conversation history.
  const requestBody = {
    model,
    messages: getConversationMessages(question)
  };

  updateDeveloperRequest(requestBody);

  addTrace("STEP 2", "REQUEST", "通常Chatリクエストを組み立て", {
    method: "POST",
    endpoint: API_ENDPOINT,
    headers: {
      Authorization: `Bearer ${maskApiKey(apiKey)}`,
      "Content-Type": "application/json; charset=utf-8",
      "X-OrcaRouter-Include-Cost": "true"
    },
    historyTurns: conversationHistory.length,
    body: requestBody
  });

  const result = await postJson(
    apiKey,
    requestBody,
    startedAt,
    "OrcaRouterへPOSTを送信",
    "Raw JSON - Chat"
  );

  const assistantText = extractAssistantText(result.json);
  const usage = result.json.usage ?? null;

  addTrace("STEP 5", "LOCAL", "Assistantメッセージを解析", {
    answerLength: assistantText.length,
    historyTurnsSent: conversationHistory.length,
    usage: usage ?? "(usage not returned)"
  });

  return {
    assistantText,
    usage,
    request: requestBody,
    response: result.json,
    status: result.response.status,
    actualModel: result.json.model ?? model
  };
}

async function runStreaming(apiKey, model, question, startedAt) {
  // STEP 2: Build request.
  const requestBody = {
    model,
    messages: getConversationMessages(question),
    stream: true,
    stream_options: { include_usage: true }
  };

  updateDeveloperRequest(requestBody);

  addTrace("STEP 2", "REQUEST", "Streamingリクエストを組み立て", {
    method: "POST",
    endpoint: API_ENDPOINT,
    historyTurns: conversationHistory.length,
    body: requestBody,
    note: "OpenAI-compatible SSE: data: {...}, terminal data: [DONE]"
  });

  addTrace("STEP 3", "REQUEST", "Streaming POSTを送信", {
    timeoutMs: STREAM_TIMEOUT_MS,
    transport: "Browser fetch + ReadableStream",
    encoding: "TextDecoder(utf-8)"
  });

  const { controller, timeoutId } = createAbortController(STREAM_TIMEOUT_MS);

  try {
    const response = await fetch(API_ENDPOINT, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json; charset=utf-8",
        Accept: "text/event-stream",
        "X-OrcaRouter-Include-Cost": "true"
      },
      body: JSON.stringify(requestBody),
      signal: controller.signal
    });

    const headers = Object.fromEntries(response.headers.entries());

    if (!response.ok) {
      const rawBody = await response.text();
      displayRawResponse(rawBody, "Raw JSON - Streaming error", response.status);
      updateDeveloperResponse(rawBody);
      addTrace("STEP 4", "RESPONSE", "Streaming開始前にHTTPエラー", {
        status: response.status,
        headers,
        rawBody
      });
      const details = buildErrorDetails(
        response.status,
        response.statusText,
        headers,
        rawBody
      );
      const error = new Error(
        `HTTP ${response.status}: ${details.errorMessage}\n${details.guidance}`
      );
      error.details = details;
      throw error;
    }

    if (!response.body) {
      throw new Error("ReadableStream が利用できません。");
    }

    addTrace("STEP 4", "RESPONSE", "SSEストリームを開始", {
      status: response.status,
      contentType: headers["content-type"] ?? "(unknown)",
      elapsedMs: Math.round(performance.now() - startedAt)
    });

    const reader = response.body.getReader();
    const decoder = new TextDecoder("utf-8");
    let buffer = "";
    let answer = "";
    let eventCount = 0;
    let usage = null;
    let latestPayload = "";
    let latestChunk = null;

    rawJsonTitle.textContent = "Raw JSON - Streaming (latest SSE event)";
    rawJsonStatus.textContent = `HTTP Status: ${response.status} / Streaming`;

    while (true) {
      const { value, done } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split(/\r?\n/);
      buffer = lines.pop() ?? "";

      for (const line of lines) {
        if (!line.startsWith("data:")) continue;

        const payload = line.slice(5).trim();

        if (!payload) continue;
        if (payload === "[DONE]") {
          addTrace("STEP 4", "STREAM", "SSE終了 [DONE]", {
            events: eventCount
          });
          continue;
        }

        latestPayload = payload;
        displayRawResponse(
          latestPayload,
          "Raw JSON - Streaming (latest SSE event)",
          response.status
        );

        let chunk;

        try {
          chunk = JSON.parse(payload);
          latestChunk = chunk;
        } catch {
          addTrace("STEP 4", "STREAM", "JSON化できないSSE data", payload);
          continue;
        }

        eventCount += 1;

        if (eventCount <= MAX_STREAM_TRACE_EVENTS) {
          addTrace("STEP 4", "STREAM", `SSE data #${eventCount}`, chunk);
        } else if (eventCount === MAX_STREAM_TRACE_EVENTS + 1) {
          addTrace(
            "STEP 4",
            "STREAM",
            "SSE Traceを省略",
            `可読性のため ${MAX_STREAM_TRACE_EVENTS} 件以降の個別イベント表示を省略します。`
          );
        }

        if (chunk.error) {
          const error = new Error(
            `Streaming error: ${chunk.error.message ?? "unknown error"}`
          );
          error.details = {
            httpStatus: response.status,
            httpStatusText: response.statusText,
            errorType: chunk.error.type ?? "",
            errorCode: chunk.error.code ?? "",
            metadata: chunk.error.metadata ?? null,
            headers,
            responseBody: chunk
          };
          throw error;
        }

        const delta = chunk?.choices?.[0]?.delta?.content;

        if (typeof delta === "string" && delta.length > 0) {
          answer += delta;
          transientAssistantText = answer;
          renderConversation(question, answer);
        }

        if (chunk.usage) {
          usage = chunk.usage;
        }
      }
    }

    buffer += decoder.decode();

    if (buffer.trim()) {
      const finalLine = buffer.trim();
      if (finalLine.startsWith("data:")) {
        const payload = finalLine.slice(5).trim();
        if (payload && payload !== "[DONE]") {
          latestPayload = payload;
          displayRawResponse(
            latestPayload,
            "Raw JSON - Streaming (latest SSE event)",
            response.status
          );

          try {
            const chunk = JSON.parse(payload);
            latestChunk = chunk;
            const delta = chunk?.choices?.[0]?.delta?.content;
            if (typeof delta === "string" && delta.length > 0) {
              answer += delta;
              transientAssistantText = answer;
              renderConversation(question, answer);
            }
            if (chunk.usage) usage = chunk.usage;
          } catch {
            addTrace("STEP 4", "STREAM", "最後のSSE dataをJSON化できません", payload);
          }
        }
      }
    }

    if (!answer) {
      throw new Error(
        "Streamingは完了しましたが回答本文を取得できませんでした。Raw JSON / Trace を確認してください。"
      );
    }

    addTrace("STEP 5", "LOCAL", "Streaming結果を集約", {
      answerLength: answer.length,
      historyTurnsSent: conversationHistory.length,
      usage: usage ?? "(usage not returned)"
    });

    return {
      assistantText: answer,
      usage,
      request: requestBody,
      response: {
        stream: true,
        latest_event: latestChunk,
        usage,
        note: "Streaming responses are SSE; latest_event is the final parsed event observed by this sample."
      },
      status: response.status,
      actualModel: latestChunk?.model ?? model
    };
  } finally {
    window.clearTimeout(timeoutId);
  }
}

function calculateSumTool(args) {
  const a = Number(args?.a);
  const b = Number(args?.b);

  if (!Number.isFinite(a) || !Number.isFinite(b)) {
    throw new Error("calculate_sum の a / b は数値である必要があります。");
  }

  return {
    a,
    b,
    sum: a + b
  };
}

async function runToolCalling(apiKey, model, question, startedAt) {
  const tools = [
    {
      type: "function",
      function: {
        name: "calculate_sum",
        description: "Add two numbers and return the sum.",
        parameters: {
          type: "object",
          properties: {
            a: { type: "number", description: "First number" },
            b: { type: "number", description: "Second number" }
          },
          required: ["a", "b"],
          additionalProperties: false
        }
      }
    }
  ];

  const firstRequest = {
    model,
    messages: getConversationMessages(question),
    tools,
    tool_choice: {
      type: "function",
      function: { name: "calculate_sum" }
    }
  };

  updateDeveloperRequest({ request_1: firstRequest });

  addTrace("STEP 2", "REQUEST", "Tool Calling 1回目のリクエストを組み立て", {
    endpoint: API_ENDPOINT,
    historyTurns: conversationHistory.length,
    body: firstRequest
  });

  const first = await postJson(
    apiKey,
    firstRequest,
    startedAt,
    "Tool Calling 1回目を送信",
    "Raw JSON - Tool Calling response #1"
  );

  const assistantMessage = first.json?.choices?.[0]?.message;
  const toolCalls = assistantMessage?.tool_calls;

  if (!Array.isArray(toolCalls) || toolCalls.length === 0) {
    throw new Error(
      "tool_calls が返されませんでした。指定モデルがTool Callingに対応しているか確認してください。"
    );
  }

  addTrace("STEP 5A", "TOOL", "モデルからTool Callを受信", toolCalls);

  const toolMessages = [];

  for (const toolCall of toolCalls) {
    if (toolCall?.function?.name !== "calculate_sum") {
      throw new Error(`未対応のToolが要求されました: ${toolCall?.function?.name}`);
    }

    let args;

    try {
      args = JSON.parse(toolCall.function.arguments || "{}");
    } catch {
      throw new Error(
        `Tool arguments JSONを解析できません: ${toolCall.function.arguments}`
      );
    }

    const result = calculateSumTool(args);

    addTrace("STEP 5B", "TOOL", "ローカル関数 calculate_sum を実行", {
      toolCallId: toolCall.id,
      arguments: args,
      result
    });

    toolMessages.push({
      role: "tool",
      tool_call_id: toolCall.id,
      content: JSON.stringify(result)
    });
  }

  const secondRequest = {
    model,
    messages: [
      ...getConversationMessages(question),
      {
        role: "assistant",
        content: assistantMessage.content ?? null,
        tool_calls: toolCalls
      },
      ...toolMessages
    ]
  };

  updateDeveloperRequest({
    request_1: firstRequest,
    request_2: secondRequest
  });

  addTrace("STEP 5C", "REQUEST", "Tool結果を含む2回目のリクエストを組み立て", {
    body: secondRequest
  });

  const second = await postJson(
    apiKey,
    secondRequest,
    startedAt,
    "Tool Calling 2回目を送信",
    "Raw JSON - Tool Calling response #2 (final)"
  );

  const assistantText = extractAssistantText(second.json);
  const usage = combineUsage(first.json.usage, second.json.usage);

  addTrace("STEP 5", "LOCAL", "Tool Calling後の最終回答を解析", {
    answerLength: assistantText.length,
    historyTurnsSent: conversationHistory.length,
    usage: usage ?? "(usage not returned)"
  });

  return {
    assistantText,
    usage,
    request: {
      request_1: firstRequest,
      request_2: secondRequest
    },
    response: {
      response_1: first.json,
      response_2: second.json
    },
    status: second.response.status,
    actualModel: second.json.model ?? model
  };
}

async function sendChat() {
  const apiKey = apiKeyInput.value.trim();
  const model = modelInput.value.trim();
  const question = questionInput.value.trim();
  const mode = modeInput.value;

  clearTracePlaceholder();
  prepareRawResponse(
    mode === "stream"
      ? "Raw JSON - Streaming"
      : mode === "tools"
        ? "Raw JSON - Tool Calling"
        : "Raw JSON - Chat"
  );
  setStatus("送信中...");
  sendButton.disabled = true;
  newChatButton.disabled = true;
  promptExampleInput.disabled = true;
  applyPromptExampleButton.disabled = true;
  transientAssistantText = "";

  const startedAt = performance.now();

  try {
    validateInputs(apiKey, model, question, mode);

    // Permanent UI contract:
    // while waiting, keep the committed conversation unchanged.
    // Do not flash the submitted question in the result area before a response exists.
    renderConversation();

    let result;

    if (mode === "stream") {
      result = await runStreaming(apiKey, model, question, startedAt);
    } else if (mode === "tools") {
      result = await runToolCalling(apiKey, model, question, startedAt);
    } else {
      result = await runChat(apiKey, model, question, startedAt);
    }

    addConversationTurn(question, result.assistantText);
    transientAssistantText = "";

    const elapsedMs = performance.now() - startedAt;
    updateDeveloperInfo({
      status: result.status,
      elapsedMs,
      model: result.actualModel || model,
      usage: result.usage,
      request: result.request,
      response: result.response
    });

    addTrace("STEP 6", "LOCAL", "画面へ回答を表示", {
      mode,
      completed: true,
      historyTurnsKept: conversationHistory.length,
      historyLimit: MAX_HISTORY_TURNS,
      totalElapsedMs: Math.round(elapsedMs)
    });

    setStatus("Completed");
  } catch (error) {
    const isAbort = error?.name === "AbortError";
    const timeoutMs = mode === "stream" ? STREAM_TIMEOUT_MS : CHAT_TOOL_TIMEOUT_MS;
    const message = isAbort
      ? `タイムアウトしました（${timeoutMs / 1000}秒）。`
      : error?.message || String(error);

    addTrace("ERROR", "ERROR", "処理中にエラーが発生", {
      message,
      name: error?.name ?? "(unknown)",
      details: error?.details ?? null,
      stack: error?.stack ?? "(stack not available)"
    });

    const errorText = transientAssistantText
      ? `${transientAssistantText}\n\n[ERROR]\n${message}`
      : `ERROR: ${message}`;

    // Keep the submitted question and the error visible without committing
    // the failed turn to conversation history.
    renderConversation(question, errorText);

    let requestForDisplay = {};
    try {
      requestForDisplay = JSON.parse(devRequestJson.textContent || "{}");
    } catch {
      requestForDisplay = { note: "Request JSON could not be parsed for display." };
    }

    updateDeveloperInfo({
      status: error?.details?.httpStatus ?? null,
      elapsedMs: performance.now() - startedAt,
      model,
      usage: null,
      request: requestForDisplay,
      response: error?.details ?? { error: message }
    });

    setStatus("Error - 会話欄とDeveloper / Traceを確認してください", true);
  } finally {
    sendButton.disabled = false;
    newChatButton.disabled = false;
    promptExampleInput.disabled = false;
    applyPromptExampleButton.disabled = false;
  }
}

modeInput.addEventListener("change", () => {
  // Changing Mode must never destroy a question the user is editing.
  setStatus(`Mode: ${modeInput.options[modeInput.selectedIndex]?.text ?? modeInput.value}`);
});

apiKeyInput.addEventListener("input", updateFirstRunHelp);

promptExampleInput.addEventListener("change", () => {
  if (!promptExampleInput.value) {
    setStatus("Ready");
    return;
  }

  const label =
    promptExampleInput.options[promptExampleInput.selectedIndex]?.text ??
    promptExampleInput.value;

  setStatus(`プロンプト例「${label}」を選択しました。［質問欄に挿入］を押してください。`);
});

applyPromptExampleButton.addEventListener("click", () => {
  const prompt = PROMPT_EXAMPLES[promptExampleInput.value];

  if (!prompt) {
    setStatus("プロンプト例を選択してください。", true);
    return;
  }

  questionInput.value = prompt;
  questionInput.focus();
  questionInput.setSelectionRange(questionInput.value.length, questionInput.value.length);
  setStatus("プロンプト例を質問欄に挿入しました。内容を編集してから送信してください。");
});

newChatButton.addEventListener("click", startNewChat);

apiKeyFileInput.addEventListener("change", async () => {
  const file = apiKeyFileInput.files?.[0];

  try {
    await loadApiKeyFromFile(file);
  } catch (error) {
    const message = error?.message || String(error);

    clearTracePlaceholder();
    addTrace("KEY", "ERROR", "APIキーファイルの読み込みに失敗", {
      source: file ? "(local file selected)" : "(not selected)",
      message
    });

    setStatus(message, true);
  } finally {
    apiKeyFileInput.value = "";
  }
});

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
updateFirstRunHelp();
updateHistoryStatus();
renderConversation();
resetDeveloperInfo();
