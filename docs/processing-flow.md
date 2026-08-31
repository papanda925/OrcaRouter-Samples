# Common processing flow

All three samples in this repository use the same six processing steps so that learners can compare the implementations side by side.

## API contract

- Endpoint: `POST https://api.orcarouter.ai/v1/chat/completions`
- Authentication: `Authorization: Bearer <API_KEY>`
- Content-Type: `application/json`
- Default model in the samples: `orcarouter/free`
- API key placeholder: `xxx-your-orcarouter-api-key-xxx`

Minimal request:

```json
{
  "model": "orcarouter/free",
  "messages": [
    {
      "role": "user",
      "content": "こんにちは"
    }
  ]
}
```

The assistant reply is read from the OpenAI-compatible response shape:

```text
choices[0].message.content
```

## Six common steps

```mermaid
sequenceDiagram
    participant U as User
    participant UI as Sample UI
    participant C as Client code
    participant O as OrcaRouter API

    U->>UI: Enter API key / model / question
    UI->>C: STEP 1 Validate inputs
    C->>C: STEP 2 Build request
    C->>O: STEP 3 Send HTTP POST
    O-->>C: STEP 4 Receive HTTP response
    C->>C: STEP 5 Parse assistant message
    C-->>UI: STEP 6 Update answer and trace
```

### STEP 1 - Validate inputs

Check the API key, model name, and question before making an HTTP request.

### STEP 2 - Build request

Create the JSON request body and HTTP headers. The trace always masks the API key.

### STEP 3 - Send HTTP POST

Send the request to OrcaRouter and record the start time.

### STEP 4 - Receive response

Capture the HTTP status, elapsed time, response headers where available, and raw response body.

### STEP 5 - Parse assistant message

Read `choices[0].message.content`. If the expected data is missing, raise an understandable error and keep the raw response in the trace.

### STEP 6 - Update UI and trace

Display the assistant answer and add the final processing information to the trace.

## Error and debug policy

The samples intentionally keep more diagnostics than production code normally would.

Each implementation records as much of the following as possible:

- timestamp
- processing step
- direction (`LOCAL`, `REQUEST`, `RESPONSE`, `ERROR`)
- endpoint and model
- masked Authorization header
- request JSON
- HTTP status
- elapsed time
- response headers when available
- raw response body
- exception/error message
- stack trace or runtime-specific error information when available

**Never write a real API key to the trace.**

These samples are designed for learning. Before using the same approach in production, add secret management, retry/backoff, telemetry rules, and application-specific security controls.
