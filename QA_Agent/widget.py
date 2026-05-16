"""
Returns the QA Agent button + modal as a single HTML string to be
injected just before </body> in every dashboard HTML file.
"""

QA_AGENT_URL = "http://localhost:5050/ask"


def qa_widget() -> str:
    return f"""
<!-- ══════════════ QA AGENT WIDGET ══════════════ -->
<style>
/* ── Floating button ── */
#qa-fab {{
  position: fixed;
  top: 16px;
  right: 22px;
  z-index: 9000;
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 9px 16px 9px 12px;
  background: linear-gradient(135deg, #3B6FD4 0%, #6C3FC5 100%);
  color: #fff;
  border: none;
  border-radius: 999px;
  font-family: inherit;
  font-size: .8rem;
  font-weight: 600;
  letter-spacing: .3px;
  cursor: pointer;
  box-shadow: 0 4px 18px rgba(60,111,212,.45);
  transition: transform .18s, box-shadow .18s;
}}
#qa-fab:hover {{
  transform: translateY(-2px);
  box-shadow: 0 7px 26px rgba(60,111,212,.55);
}}
#qa-fab svg {{ flex-shrink: 0; }}
#qa-fab .qa-pulse {{
  width: 7px; height: 7px;
  background: #4ade80;
  border-radius: 50%;
  animation: qa-blink 2s ease-in-out infinite;
}}
@keyframes qa-blink {{ 0%,100%{{opacity:1}} 50%{{opacity:.35}} }}

/* ── Overlay ── */
#qa-overlay {{
  display: none;
  position: fixed;
  inset: 0;
  z-index: 9100;
  background: rgba(8,12,22,.72);
  backdrop-filter: blur(5px);
  -webkit-backdrop-filter: blur(5px);
  align-items: center;
  justify-content: center;
}}
#qa-overlay.qa-open {{ display: flex; }}

/* ── Modal ── */
#qa-modal {{
  width: min(660px, 96vw);
  height: min(580px, 88vh);
  background: #111827;
  border: 1px solid rgba(99,130,220,.3);
  border-radius: 18px;
  box-shadow: 0 30px 80px rgba(0,0,0,.7);
  display: flex;
  flex-direction: column;
  overflow: hidden;
  animation: qa-slide-in .22s cubic-bezier(.16,1,.3,1);
}}
@keyframes qa-slide-in {{
  from {{ opacity:0; transform: scale(.94) translateY(14px); }}
  to   {{ opacity:1; transform: scale(1)  translateY(0); }}
}}

/* ── Modal header ── */
#qa-header {{
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 16px 20px 14px;
  border-bottom: 1px solid rgba(99,130,220,.2);
  background: linear-gradient(90deg, rgba(59,111,212,.18) 0%, rgba(108,63,197,.12) 100%);
  flex-shrink: 0;
}}
#qa-header-left {{
  display: flex;
  align-items: center;
  gap: 10px;
}}
#qa-header-icon {{
  width: 34px; height: 34px;
  background: linear-gradient(135deg, #3B6FD4, #6C3FC5);
  border-radius: 10px;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
}}
#qa-header-title {{ font-size: .95rem; font-weight: 700; color: #e8edf8; }}
#qa-header-sub   {{ font-size: .7rem;  color: #8896b4; margin-top: 1px; }}
#qa-close-btn {{
  background: rgba(255,255,255,.07);
  border: 1px solid rgba(255,255,255,.1);
  color: #8896b4;
  width: 30px; height: 30px;
  border-radius: 8px;
  cursor: pointer;
  display: flex; align-items: center; justify-content: center;
  transition: background .15s, color .15s;
  font-size: 1.1rem; line-height: 1;
}}
#qa-close-btn:hover {{ background: rgba(255,80,80,.18); color: #ef9a9a; }}

/* ── Suggestions bar ── */
#qa-suggestions {{
  display: flex;
  gap: 7px;
  padding: 10px 16px 8px;
  overflow-x: auto;
  flex-shrink: 0;
  scrollbar-width: none;
  border-bottom: 1px solid rgba(99,130,220,.12);
}}
#qa-suggestions::-webkit-scrollbar {{ display: none; }}
.qa-chip {{
  flex-shrink: 0;
  padding: 5px 11px;
  background: rgba(59,111,212,.14);
  border: 1px solid rgba(59,111,212,.28);
  border-radius: 999px;
  font-size: .72rem;
  color: #93b4ee;
  cursor: pointer;
  white-space: nowrap;
  transition: background .15s, color .15s;
}}
.qa-chip:hover {{ background: rgba(59,111,212,.3); color: #cce0ff; }}

/* ── Messages ── */
#qa-messages {{
  flex: 1;
  overflow-y: auto;
  padding: 14px 18px;
  display: flex;
  flex-direction: column;
  gap: 14px;
  scrollbar-width: thin;
  scrollbar-color: rgba(99,130,220,.3) transparent;
}}
.qa-msg {{ display: flex; gap: 10px; align-items: flex-start; max-width: 100%; }}
.qa-msg-avatar {{
  width: 28px; height: 28px;
  border-radius: 8px;
  flex-shrink: 0;
  display: flex; align-items: center; justify-content: center;
  font-size: .75rem; font-weight: 700;
}}
.qa-msg-user .qa-msg-avatar  {{ background: rgba(59,111,212,.3); color: #93b4ee; }}
.qa-msg-agent .qa-msg-avatar {{ background: linear-gradient(135deg,#3B6FD4,#6C3FC5); color:#fff; }}
.qa-msg-bubble {{
  padding: 10px 13px;
  border-radius: 12px;
  font-size: .82rem;
  line-height: 1.55;
  max-width: calc(100% - 40px);
  word-wrap: break-word;
}}
.qa-msg-user .qa-msg-bubble {{
  background: rgba(59,111,212,.18);
  border: 1px solid rgba(59,111,212,.25);
  color: #d4e0f8;
  border-top-left-radius: 4px;
}}
.qa-msg-agent .qa-msg-bubble {{
  background: rgba(30,38,60,.8);
  border: 1px solid rgba(99,130,220,.2);
  color: #cbd5ef;
  border-top-left-radius: 4px;
}}
.qa-msg-bubble strong {{ color: #e8edf8; }}
.qa-msg-bubble em     {{ color: #93b4ee; font-style: normal; }}
.qa-typing {{
  display: flex; gap: 5px; align-items: center; padding: 6px 4px;
}}
.qa-dot {{
  width: 7px; height: 7px;
  background: #5878c8;
  border-radius: 50%;
  animation: qa-bounce .9s ease-in-out infinite;
}}
.qa-dot:nth-child(2) {{ animation-delay: .15s; }}
.qa-dot:nth-child(3) {{ animation-delay: .3s; }}
@keyframes qa-bounce {{ 0%,80%,100%{{transform:translateY(0)}} 40%{{transform:translateY(-6px)}} }}

/* ── Welcome message ── */
#qa-welcome {{
  text-align: center;
  padding: 28px 20px;
  color: #6878a8;
  font-size: .8rem;
  line-height: 1.6;
}}
#qa-welcome svg {{ opacity: .35; margin-bottom: 12px; }}
#qa-welcome strong {{ display: block; color: #8896c8; font-size: .9rem; margin-bottom: 6px; }}

/* ── Input area ── */
#qa-input-area {{
  display: flex;
  gap: 10px;
  padding: 12px 16px;
  border-top: 1px solid rgba(99,130,220,.2);
  background: rgba(15,20,40,.6);
  flex-shrink: 0;
  align-items: flex-end;
}}
#qa-input {{
  flex: 1;
  background: rgba(255,255,255,.06);
  border: 1px solid rgba(99,130,220,.28);
  border-radius: 10px;
  padding: 9px 13px;
  color: #e8edf8;
  font-family: inherit;
  font-size: .83rem;
  outline: none;
  resize: none;
  min-height: 38px;
  max-height: 100px;
  line-height: 1.4;
  transition: border-color .15s;
}}
#qa-input::placeholder {{ color: #4a5878; }}
#qa-input:focus {{ border-color: rgba(99,130,220,.6); }}
#qa-send-btn {{
  flex-shrink: 0;
  width: 38px; height: 38px;
  background: linear-gradient(135deg, #3B6FD4, #6C3FC5);
  border: none;
  border-radius: 10px;
  color: #fff;
  cursor: pointer;
  display: flex; align-items: center; justify-content: center;
  transition: opacity .15s, transform .15s;
}}
#qa-send-btn:hover {{ opacity: .85; transform: scale(1.06); }}
#qa-send-btn:disabled {{ opacity: .35; cursor: not-allowed; transform: none; }}
#qa-footer-note {{
  padding: 6px 16px 10px;
  font-size: .63rem;
  color: #3a4668;
  text-align: center;
  flex-shrink: 0;
}}
</style>

<!-- Floating button -->
<button id="qa-fab" onclick="qaToggle()">
  <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
  </svg>
  Ask the Data
  <span class="qa-pulse"></span>
</button>

<!-- Overlay + Modal -->
<div id="qa-overlay" onclick="if(event.target===this)qaToggle()">
  <div id="qa-modal">

    <!-- Header -->
    <div id="qa-header">
      <div id="qa-header-left">
        <div id="qa-header-icon">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
          </svg>
        </div>
        <div>
          <div id="qa-header-title">HPI Data Assistant</div>
          <div id="qa-header-sub">Ask questions about the UK House Price Index data</div>
        </div>
      </div>
      <button id="qa-close-btn" onclick="qaToggle()" title="Close">&#x2715;</button>
    </div>

    <!-- Suggestion chips -->
    <div id="qa-suggestions">
      <span class="qa-chip" onclick="qaAsk(this.textContent)">Which region grew fastest?</span>
      <span class="qa-chip" onclick="qaAsk(this.textContent)">Average price in 2010?</span>
      <span class="qa-chip" onclick="qaAsk(this.textContent)">Most expensive region?</span>
      <span class="qa-chip" onclick="qaAsk(this.textContent)">Detached vs flat prices?</span>
      <span class="qa-chip" onclick="qaAsk(this.textContent)">Sales volume trends?</span>
      <span class="qa-chip" onclick="qaAsk(this.textContent)">YoY growth peak?</span>
    </div>

    <!-- Message thread -->
    <div id="qa-messages">
      <div id="qa-welcome">
        <svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
          <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
        </svg>
        <strong>Ask me anything about the data</strong>
        6 English metropolitan regions &nbsp;·&nbsp; Jan 1995 – Oct 2016<br/>
        Prices, volumes, growth rates, property types &amp; more
      </div>
    </div>

    <!-- Input -->
    <div id="qa-input-area">
      <textarea id="qa-input" placeholder="e.g. Which region had the highest price growth?" rows="1"
        onkeydown="if(event.key==='Enter'&&!event.shiftKey){{event.preventDefault();qaSend();}}"
        oninput="this.style.height='auto';this.style.height=Math.min(this.scrollHeight,100)+'px'"></textarea>
      <button id="qa-send-btn" onclick="qaSend()" title="Send">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round">
          <line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/>
        </svg>
      </button>
    </div>
    <div id="qa-footer-note">Powered by Claude · Answers are based on dashboard data only · Requires QA Agent server running locally</div>

  </div>
</div>

<script>
(function() {{
  const QA_URL = "{QA_AGENT_URL}";
  let sessionId = null;
  let busy = false;

  window.qaToggle = function() {{
    const ov = document.getElementById('qa-overlay');
    ov.classList.toggle('qa-open');
    if (ov.classList.contains('qa-open')) {{
      document.getElementById('qa-input').focus();
    }}
  }};

  window.qaAsk = function(text) {{
    document.getElementById('qa-input').value = text;
    qaSend();
  }};

  window.qaSend = function() {{
    const input = document.getElementById('qa-input');
    const question = input.value.trim();
    if (!question || busy) return;

    appendMsg('user', question);
    input.value = '';
    input.style.height = 'auto';

    const typingId = appendTyping();
    busy = true;
    document.getElementById('qa-send-btn').disabled = true;

    fetch(QA_URL, {{
      method: 'POST',
      headers: {{'Content-Type': 'application/json'}},
      body: JSON.stringify({{ question, session_id: sessionId }})
    }})
    .then(r => r.json())
    .then(data => {{
      removeTyping(typingId);
      if (data.error) {{
        appendMsg('agent', '⚠ ' + data.error);
      }} else {{
        sessionId = data.session_id;
        appendMsg('agent', data.answer);
      }}
    }})
    .catch(() => {{
      removeTyping(typingId);
      appendMsg('agent',
        '⚠ Could not reach the QA Agent server.\\n\\n' +
        'Make sure the server is running:\\n' +
        '  cd QA_Agent && .\\\\start.ps1'
      );
    }})
    .finally(() => {{
      busy = false;
      document.getElementById('qa-send-btn').disabled = false;
      document.getElementById('qa-input').focus();
    }});
  }};

  function appendMsg(role, text) {{
    const welcome = document.getElementById('qa-welcome');
    if (welcome) welcome.remove();

    const container = document.getElementById('qa-messages');
    const div = document.createElement('div');
    div.className = 'qa-msg qa-msg-' + role;

    const avatar = document.createElement('div');
    avatar.className = 'qa-msg-avatar';
    avatar.textContent = role === 'user' ? 'You' : 'AI';

    const bubble = document.createElement('div');
    bubble.className = 'qa-msg-bubble';
    bubble.innerHTML = formatText(text);

    div.appendChild(avatar);
    div.appendChild(bubble);
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
    return div;
  }}

  function appendTyping() {{
    const container = document.getElementById('qa-messages');
    const div = document.createElement('div');
    div.className = 'qa-msg qa-msg-agent';
    div.id = 'qa-typing-' + Date.now();

    const avatar = document.createElement('div');
    avatar.className = 'qa-msg-avatar';
    avatar.textContent = 'AI';

    const bubble = document.createElement('div');
    bubble.className = 'qa-msg-bubble';
    bubble.innerHTML = '<div class="qa-typing"><div class="qa-dot"></div><div class="qa-dot"></div><div class="qa-dot"></div></div>';

    div.appendChild(avatar);
    div.appendChild(bubble);
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
    return div.id;
  }}

  function removeTyping(id) {{
    const el = document.getElementById(id);
    if (el) el.remove();
  }}

  function formatText(text) {{
    return text
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/\\*\\*(.+?)\\*\\*/g, '<strong>$1</strong>')
      .replace(/`(.+?)`/g, '<em>$1</em>')
      .replace(/\\n/g, '<br/>');
  }}
}})();
</script>
<!-- ══════════════ END QA AGENT WIDGET ══════════════ -->
"""
