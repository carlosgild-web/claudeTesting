import uuid
from flask import Flask, request, jsonify
from flask_cors import CORS
from agent import ask, get_system_prompt

app = Flask(__name__)
CORS(app)  # allow calls from local file:// dashboards

# In-memory session store: session_id -> list of {role, content}
_sessions: dict[str, list] = {}


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/ask", methods=["POST"])
def handle_ask():
    data = request.get_json(silent=True) or {}
    question   = (data.get("question") or "").strip()
    session_id = data.get("session_id") or str(uuid.uuid4())

    if not question:
        return jsonify({"error": "No question provided"}), 400

    history = _sessions.get(session_id, [])

    try:
        answer = ask(question, history)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

    # Persist conversation (cap at 40 messages = 20 turns)
    history.append({"role": "user",      "content": question})
    history.append({"role": "assistant", "content": answer})
    _sessions[session_id] = history[-40:]

    return jsonify({"answer": answer, "session_id": session_id})


@app.route("/reset", methods=["POST"])
def reset_session():
    data = request.get_json(silent=True) or {}
    sid  = data.get("session_id", "")
    _sessions.pop(sid, None)
    return jsonify({"status": "reset"})


if __name__ == "__main__":
    print("Warming up QA Agent — loading data context…")
    get_system_prompt()  # pre-load so first query is fast
    print("QA Agent is ready.")
    print("Listening on http://localhost:5050")
    print("Press Ctrl+C to stop.\n")
    app.run(host="127.0.0.1", port=5050, debug=False)
