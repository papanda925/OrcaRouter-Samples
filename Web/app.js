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

apiKeyInput.value = DEFAULT_API_KEY;

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
    guidance
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
        Accept: "application/json"
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

    return {
      response,
      headers,
      rawBody,
      json: rawBody ? JSON.parse(rawBody) : {}
    };
  } finally {
    window.clearTimeout(timeoutId);
  }
}

async function runChat(apiKey, model, question, startedAt) {
  // STEP 2: Build request.
  const requestBody = {
    model,
    messages: [{ role: "user", content: question }]
  };

  addTrace("STEP 2", "REQUEST", "通常Chatリクエストを組み立て", {
    method: "POST",
    endpoint: API_ENDPOINT,
    headers: {
      Authorization: `Bearer ${maskApiKey(apiKey)}`,
      "Content-Type": "application/json; charset=utf-8"
    },
    body: requestBody
  });

  const result = await postJson(
    apiKey,
    requestBody,
    startedAt,
    "OrcaRouterへPOSTを送信",
    "Raw JSON - Chat"
  );

  // STEP 5: Parse assistant message.
  const assistantText = extractAssistantText(result.json);

  addTrace("STEP 5", "LOCAL", "Assistantメッセージを解析", {
    answerLength: assistantText.length,
    usage: result.json.usage ?? "(usage not returned)"
  });

  return assistantText;
}

async function runStreaming(apiKey, model, question, startedAt) {
  // STEP 2: Build request.
  const requestBody = {
    model,
    messages: [{ role: "user", content: question }],
    stream: true,
    stream_options: { include_usage: true }
  };

  addTrace("STEP 2", "REQUEST", "Streamingリクエストを組み立て", {
    method: "POST",
    endpoint: API_ENDPOINT,
    body: requestBody,
    note: "OpenAI-compatible SSE: data: {...}, terminal data: [DONE]"
  });

  // STEP 3: Send HTTP POST.
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
        Accept: "text/event-stream"
      },
      body: JSON.stringify(requestBody),
      signal: controller.signal
    });

    const headers = Object.fromEntries(response.headers.entries());

    if (!response.ok) {
      const rawBody = await response.text();
      displayRawResponse(rawBody, "Raw JSON - Streaming error", response.status);
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
        } catch (error) {
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
            errorType: chunk.error.type ?? "",
            errorCode: chunk.error.code ?? "",
            metadata: chunk.error.metadata ?? null
          };
          throw error;
        }

        const delta = chunk?.choices?.[0]?.delta?.content;

        if (typeof delta === "string" && delta.length > 0) {
          answer += delta;
          answerBox.textContent = answer;
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
            const delta = chunk?.choices?.[0]?.delta?.content;
            if (typeof delta === "string" && delta.length > 0) {
              answer += delta;
              answerBox.textContent = answer;
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

    // STEP 5: Parse assistant message.
    addTrace("STEP 5", "LOCAL", "Streaming結果を集約", {
      answerLength: answer.length,
      usage: usage ?? "(usage not returned)"
    });

    return answer;
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

  // STEP 2: Build request.
  const firstRequest = {
    model,
    messages: [{ role: "user", content: question }],
    tools,
    tool_choice: {
      type: "function",
      function: { name: "calculate_sum" }
    }
  };

  addTrace("STEP 2", "REQUEST", "Tool Calling 1回目のリクエストを組み立て", {
    endpoint: API_ENDPOINT,
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
    } catch (error) {
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
      { role: "user", content: question },
      {
        role: "assistant",
        content: assistantMessage.content ?? null,
        tool_calls: toolCalls
      },
      ...toolMessages
    ]
  };

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

  addTrace("STEP 5", "LOCAL", "Tool Calling後の最終回答を解析", {
    answerLength: assistantText.length,
    usage: second.json.usage ?? "(usage not returned)"
  });

  return assistantText;
}

async function sendChat() {
  const apiKey = apiKeyInput.value.trim();
  const model = modelInput.value.trim();
  const question = questionInput.value.trim();
  const mode = modeInput.value;

  clearTracePlaceholder();
  answerBox.textContent = "";
  prepareRawResponse(
    mode === "stream"
      ? "Raw JSON - Streaming"
      : mode === "tools"
        ? "Raw JSON - Tool Calling"
        : "Raw JSON - Chat"
  );
  setStatus("Processing...");
  sendButton.disabled = true;

  const startedAt = performance.now();

  try {
    validateInputs(apiKey, model, question, mode);

    let assistantText;

    if (mode === "stream") {
      assistantText = await runStreaming(apiKey, model, question, startedAt);
    } else if (mode === "tools") {
      assistantText = await runToolCalling(apiKey, model, question, startedAt);
    } else {
      assistantText = await runChat(apiKey, model, question, startedAt);
    }

    // STEP 6: Update UI and trace.
    answerBox.textContent = assistantText;

    addTrace("STEP 6", "LOCAL", "画面へ回答を表示", {
      mode,
      completed: true,
      totalElapsedMs: Math.round(performance.now() - startedAt)
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

    answerBox.textContent = `ERROR: ${message}`;
    setStatus("Error - Trace を確認してください", true);
  } finally {
    sendButton.disabled = false;
  }
}

modeInput.addEventListener("change", () => {
  if (modeInput.value === "tools") {
    questionInput.value =
      "calculate_sum ツールを使って 123 と 456 を足し、その結果を日本語で説明してください。";
  } else if (modeInput.value === "stream") {
    questionInput.value =
      "日本語で「こんにちは。Streamingのテストです。」と短く答えてください。";
  } else {
    questionInput.value =
      "日本語で「こんにちは。Web版Chatのテストです。」とだけ答えてください。";
  }
});

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
