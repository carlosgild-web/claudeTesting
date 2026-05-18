"""Generate QA_Agent Production Deployment Plan PDF."""
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, Preformatted
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER
import os

OUTPUT = os.path.join(os.path.dirname(__file__), "QA_Agent_Production_Plan.pdf")

# ── Colour palette ──────────────────────────────────────────────────────────
NAVY   = colors.HexColor("#1B3A5C")
GOLD   = colors.HexColor("#C9A84C")
SILVER = colors.HexColor("#F0F0F0")
DARK   = colors.HexColor("#2D2D2D")
CODE_BG= colors.HexColor("#F4F4F4")

# ── Styles ───────────────────────────────────────────────────────────────────
base = getSampleStyleSheet()

def style(name, parent="Normal", **kw):
    s = ParagraphStyle(name, parent=base[parent], **kw)
    return s

S_TITLE    = style("S_TITLE",    "Title",   fontSize=20, textColor=NAVY,
                   spaceAfter=4, leading=24, alignment=TA_CENTER)
S_SUBTITLE = style("S_SUBTITLE", "Normal",  fontSize=10, textColor=GOLD,
                   spaceAfter=16, alignment=TA_CENTER)
S_H1       = style("S_H1",       "Heading1", fontSize=13, textColor=NAVY,
                   spaceBefore=14, spaceAfter=4, leading=16)
S_H2       = style("S_H2",       "Heading2", fontSize=11, textColor=NAVY,
                   spaceBefore=10, spaceAfter=3, leading=14)
S_BODY     = style("S_BODY",     "Normal",   fontSize=9,  textColor=DARK,
                   spaceAfter=5, leading=13)
S_BULLET   = style("S_BULLET",   "Normal",   fontSize=9,  textColor=DARK,
                   leftIndent=14, spaceAfter=3, leading=13,
                   bulletIndent=5, bulletFontName="Helvetica")
S_CODE     = style("S_CODE",     "Code",     fontSize=7.5, backColor=CODE_BG,
                   leftIndent=12, rightIndent=12, spaceBefore=4, spaceAfter=4,
                   leading=11, fontName="Courier")
S_NOTE     = style("S_NOTE",     "Normal",   fontSize=8.5, textColor=colors.HexColor("#555"),
                   leftIndent=10, spaceAfter=6, leading=12, borderPad=4)

def hr():
    return HRFlowable(width="100%", thickness=0.5, color=GOLD, spaceAfter=8, spaceBefore=4)

def h1(txt):  return Paragraph(txt, S_H1)
def h2(txt):  return Paragraph(txt, S_H2)
def body(txt): return Paragraph(txt, S_BODY)
def bullet(txt): return Paragraph(f"•  {txt}", S_BULLET)
def note(txt):  return Paragraph(f"<i>{txt}</i>", S_NOTE)
def code(txt):  return Preformatted(txt, S_CODE)
def sp(n=6):   return Spacer(1, n)

def tbl(data, col_widths, header=True):
    t = Table(data, colWidths=col_widths, repeatRows=1 if header else 0)
    style_cmds = [
        ("BACKGROUND",  (0, 0), (-1, 0 if header else -1), NAVY),
        ("TEXTCOLOR",   (0, 0), (-1, 0 if header else -1), colors.white),
        ("FONTNAME",    (0, 0), (-1, 0 if header else -1), "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, SILVER]),
        ("GRID",        (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
        ("VALIGN",      (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING",   (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
    ]
    if not header:
        style_cmds[0] = ("BACKGROUND", (0,0), (-1,-1), colors.white)
        style_cmds[1] = ("TEXTCOLOR",  (0,0), (-1,-1), DARK)
        style_cmds[2] = ("FONTNAME",   (0,0), (-1,-1), "Helvetica")
    t.setStyle(TableStyle(style_cmds))
    return t

# ── Build content ─────────────────────────────────────────────────────────────
def build():
    doc = SimpleDocTemplate(
        OUTPUT,
        pagesize=A4,
        topMargin=2*cm, bottomMargin=2*cm,
        leftMargin=2*cm, rightMargin=2*cm,
        title="QA_Agent Production Deployment Plan",
        author="HPI Dashboard Project",
    )

    story = []

    # Cover
    story += [
        sp(20),
        Paragraph("QA_Agent Production Deployment Plan", S_TITLE),
        Paragraph("Deploy QA_Agent to AWS EC2 for Organization-Wide Access", S_SUBTITLE),
        hr(),
        sp(4),
        tbl(
            [["Prepared", "HPI Dashboard Project"],
             ["Date",     "18 May 2026"],
             ["Platform", "AWS EC2 (private VPC) · Flask · Anthropic Claude API"],
             ["Security", "Corporate VPN / network boundary · AWS SSM for API key"]],
            [4*cm, 12*cm], header=False
        ),
        sp(20),
    ]

    # ── Context ──────────────────────────────────────────────────────────────
    story += [hr(), h1("Context"), sp(2)]
    story.append(body(
        "The HTML dashboards embed a QA chat widget (<b>QA_Agent/widget.py</b>) that "
        "currently calls <b>http://localhost:5050/ask</b> — a Flask server that must be "
        "running on the <i>same machine</i> as the browser. When users in the organization "
        "receive the HTML files, the QA_Agent silently fails because no server is running "
        "on their machines."
    ))
    story.append(body(
        "The fix is to deploy the existing Flask server once, centrally, on AWS EC2 inside "
        "the corporate VPC (accessible over VPN), then update one URL in the widget so all "
        "rebuilt dashboards point to the shared server. The Anthropic API key is held only "
        "on the EC2 instance — never embedded in the HTML files."
    ))

    # ── Architecture ─────────────────────────────────────────────────────────
    story += [sp(4), hr(), h1("Architecture After This Change"), sp(2)]
    story.append(code(
        "User's machine (file://*.html)\n"
        "        │\n"
        "        │  POST http://<ec2-private-ip>:5050/ask\n"
        "        ▼  (reachable only via corporate VPN / corporate network)\n"
        "AWS EC2 t3.small — private subnet\n"
        "  └── python QA_Agent/server.py  (systemd service, port 5050)\n"
        "        │  Anthropic Python SDK\n"
        "        ▼\n"
        "Anthropic Claude API  (claude-haiku-4-5-20251001)\n"
    ))
    story.append(note(
        "HTML files remain self-contained static files — no web server required to view them. "
        "The only network call they make is POST /ask to the EC2 backend."
    ))

    # ── Code changes ─────────────────────────────────────────────────────────
    story += [sp(4), hr(), h1("Code Changes Required"), sp(2)]

    story.append(h2("1.  QA_Agent/widget.py  (line ~6)"))
    story.append(body("Replace the hardcoded localhost URL with an environment-variable lookup:"))
    story.append(code(
        "# Before\n"
        'QA_URL = "http://localhost:5050/ask"\n\n'
        "# After\n"
        "import os\n"
        'QA_URL = os.environ.get("QA_AGENT_URL", "http://localhost:5050") + "/ask"'
    ))
    story.append(body(
        "Setting <b>QA_AGENT_URL</b> in the shell before running the build script bakes "
        "the correct URL into every generated dashboard."
    ))

    story.append(h2("2.  QA_Agent/server.py — CORS (already configured)"))
    story.append(body(
        "flask-cors is already enabled. Verify <b>CORS(app)</b> has no origin restriction. "
        "file:// dashboards carry no Origin header; flask-cors handles this correctly with "
        "its default settings — no change needed."
    ))

    story.append(h2("3.  New deployment files  QA_Agent/deploy/"))
    story.append(tbl(
        [["File", "Purpose"],
         ["ec2_setup.sh",        "One-shot setup: installs Python 3.11, pip deps, copies QA_Agent/"],
         ["hpi-qa-agent.service","systemd unit — auto-starts server on reboot, references API key path"],
         ["ssm_param.md",        "Instructions: store ANTHROPIC_API_KEY in AWS SSM Parameter Store"]],
        [5*cm, 11*cm]
    ))

    story.append(h2("Files to Modify Summary"))
    story.append(tbl(
        [["File", "Change"],
         ["QA_Agent/widget.py line ~6",      "Replace hardcoded URL with env-var pattern"],
         ["QA_Agent/deploy/ec2_setup.sh",    "NEW — EC2 bootstrap shell script"],
         ["QA_Agent/deploy/hpi-qa-agent.service", "NEW — systemd unit file"]],
        [7*cm, 9*cm]
    ))
    story.append(note(
        "build_dashboards_all.py and build_dashboard.py require NO changes — "
        "they call qa_widget() which reads the updated widget.py."
    ))

    # ── Deployment steps ─────────────────────────────────────────────────────
    story += [sp(4), hr(), h1("Deployment Steps"), sp(2)]

    story.append(h2("Phase 1 — AWS Infrastructure  (one-time)"))
    story.append(body("<b>Step 1: Launch EC2 instance</b>"))
    for t in [
        "AMI: Amazon Linux 2023",
        "Type: t3.small (≈ $15 / month)",
        "VPC: private subnet reachable from corporate VPN",
        "Security Group: allow TCP 5050 inbound from corporate IP range ONLY",
        "No public IP required",
    ]:
        story.append(bullet(t))

    story.append(sp(4))
    story.append(body("<b>Step 2: Store API key in AWS SSM Parameter Store</b>"))
    story.append(code(
        "aws ssm put-parameter \\\n"
        "  --name /hpi/qa-agent/anthropic-api-key \\\n"
        '  --value "sk-ant-..." \\\n'
        "  --type SecureString"
    ))
    story.append(note("The key never touches disk in plain text and never appears in the HTML files."))

    story.append(body("<b>Step 3: Run setup script on EC2</b>  (SSH or AWS Session Manager)"))
    story.append(code("bash ec2_setup.sh"))

    story.append(body("<b>Step 4: Start and enable the service</b>"))
    story.append(code(
        "sudo systemctl enable hpi-qa-agent\n"
        "sudo systemctl start  hpi-qa-agent\n"
        "curl http://localhost:5050/health   # expected: {\"status\":\"ok\"}"
    ))

    story.append(body("<b>Step 5:</b> Note the EC2 private IP address (e.g. <b>10.0.1.45</b>)."))

    story.append(h2("Phase 2 — Rebuild Dashboards with New URL"))
    story.append(code(
        "$env:QA_AGENT_URL = \"http://10.0.1.45:5050\"\n"
        "python build_dashboards_all.py"
    ))
    story.append(body(
        "The four rebuilt HTML files in <b>20260518/</b> will call "
        "http://10.0.1.45:5050/ask instead of localhost."
    ))

    story.append(h2("Phase 3 — Distribute the Rebuilt HTML Files"))
    story.append(body(
        "No change to distribution method. Share the four HTML files via email or "
        "SharePoint as before. Users connected to VPN (or the corporate network) open the "
        "file in any browser — the QA chat widget works automatically."
    ))

    # ── Security ─────────────────────────────────────────────────────────────
    story += [sp(4), hr(), h1("Security Controls"), sp(2)]
    story.append(tbl(
        [["Control", "Implementation"],
         ["Network boundary",       "EC2 Security Group: port 5050 from corporate IP range only"],
         ["API key protection",     "Stored in AWS SSM Parameter Store; injected as env var; never in HTML"],
         ["Authentication",         "VPN / corporate network boundary is the auth layer — no tokens needed"],
         ["CORS",                   "flask-cors default (permissive for file:// — no Origin header sent)"],
         ["Rate limiting (future)", "Add flask-limiter: 100 req/hour per IP if abuse is a concern"]],
        [5*cm, 11*cm]
    ))

    # ── Verification ─────────────────────────────────────────────────────────
    story += [sp(4), hr(), h1("Verification"), sp(2)]

    story.append(h2("Server-side health check  (EC2 terminal)"))
    story.append(code(
        "curl -s http://localhost:5050/health\n"
        '# expected: {"status":"ok"}\n\n'
        "curl -s -X POST http://localhost:5050/ask \\\n"
        "  -H 'Content-Type: application/json' \\\n"
        "  -d '{\"question\":\"What is the average price in Greater Manchester?\"}'\n"
        '# expected: {"answer":"...","session_id":"..."}'
    ))

    story.append(h2("End-to-end test from a VPN-connected machine"))
    for t in [
        "Open rebuilt CEO_HPIAdmins.html in a browser",
        "Click the 'Ask the Data' floating button",
        "Type a question — should receive an answer within a few seconds",
    ]:
        story.append(bullet(t))

    story.append(h2("Negative test  (outside VPN)"))
    story.append(body(
        "Opening the same HTML file from outside the VPN should display: "
        "<b>⚠ Could not reach the QA Agent server.</b> — confirming the network boundary is enforced."
    ))

    # ── User experience ───────────────────────────────────────────────────────
    story += [sp(4), hr(), h1("What Users Need To Do"), sp(2)]
    story.append(body(
        "<b>Nothing.</b>  Users open the HTML file as normal. If they are on VPN or the "
        "corporate network the QA_Agent works automatically.  No Python installation, no "
        "API key, no local server setup required on their machines."
    ))
    story.append(sp(8))
    story.append(tbl(
        [["Who", "Action required"],
         ["IT / DevOps",        "Deploy EC2 instance once (Phases 1-2 above)"],
         ["Dashboard publisher", "Set QA_AGENT_URL env var and re-run build_dashboards_all.py"],
         ["End users",          "Open HTML file as usual — QA_Agent works if on VPN"]],
        [5*cm, 11*cm]
    ))

    doc.build(story)
    print(f"PDF saved to: {OUTPUT}")

if __name__ == "__main__":
    build()
