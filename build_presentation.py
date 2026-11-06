# -*- coding: utf-8 -*-
"""
EKATVA AI — Pitch Deck Builder
Creates: EKATVA_Entrepreneurship_Presentation.pptx
15 slides, 16:9, dark-navy / electric-blue / purple premium design.
Every slide includes detailed speaker notes (Vipul: 1-8, Mohan: 9-15).
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE
import math

# ─────────────────────────── DESIGN TOKENS ───────────────────────────
SLIDE_W = 13.333
SLIDE_H = 7.5

NAVY      = RGBColor(0x0A, 0x12, 0x2E)   # base dark navy
NAVY2     = RGBColor(0x17, 0x1B, 0x4D)   # gradient partner (purple-navy)
CARD      = RGBColor(0x13, 0x1B, 0x3F)
CARD_HI   = RGBColor(0x1A, 0x2E, 0x64)
BORDER    = RGBColor(0x2C, 0x3D, 0x7A)
BLUE      = RGBColor(0x00, 0xC6, 0xFF)   # electric blue
BLUE_D    = RGBColor(0x1E, 0x6F, 0xFF)
PURPLE    = RGBColor(0x8B, 0x5C, 0xF6)
PURPLE_D  = RGBColor(0x5B, 0x3A, 0xD1)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
SOFT      = RGBColor(0xC7, 0xD2, 0xEE)   # soft body text
MUTED     = RGBColor(0x8E, 0x9B, 0xC2)   # muted text
GOLD      = RGBColor(0xFF, 0xC9, 0x50)
DANGER    = RGBColor(0xFF, 0x8A, 0x7A)

FONT = "Segoe UI"

prs = Presentation()
prs.slide_width  = Inches(SLIDE_W)
prs.slide_height = Inches(SLIDE_H)
BLANK = prs.slide_layouts[6]

# ─────────────────────────── HELPERS ───────────────────────────

def no_shadow(shape):
    try:
        shape.shadow.inherit = False
    except Exception:
        pass

def new_slide():
    return prs.slides.add_slide(BLANK)

def bg(slide, angle=115.0):
    """Full-slide navy->purple gradient background."""
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0,
                                Inches(SLIDE_W), Inches(SLIDE_H))
    no_shadow(sh)
    f = sh.fill
    f.gradient()
    stops = f.gradient_stops
    stops[0].color.rgb = NAVY
    stops[0].position = 0.0
    stops[1].color.rgb = NAVY2
    stops[1].position = 1.0
    try:
        f.gradient_angle = angle
    except Exception:
        pass
    sh.line.fill.background()
    # small accent glow line bottom
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0),
                                Inches(SLIDE_H - 0.06),
                                Inches(SLIDE_W), Inches(0.06))
    no_shadow(ln)
    lf = ln.fill
    lf.gradient()
    lstops = lf.gradient_stops
    lstops[0].color.rgb = BLUE
    lstops[1].color.rgb = PURPLE
    ln.line.fill.background()

def tx(slide, x, y, w, h, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
       wrap=True):
    """paras: list of dicts(text,size,color,bold,italic,space_before,space_after,align,line)"""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    for i, d in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = d.get("align", align)
        if d.get("space_before"):
            p.space_before = Pt(d["space_before"])
        p.space_after = Pt(d.get("space_after", 2))
        if d.get("line"):
            p.line_spacing = d["line"]
        r = p.add_run()
        r.text = d["text"]
        r.font.name = FONT
        r.font.size = Pt(d.get("size", 12))
        r.font.color.rgb = d.get("color", WHITE)
        r.font.bold = d.get("bold", False)
        r.font.italic = d.get("italic", False)
    return box

def card(slide, x, y, w, h, fill=CARD, line=BORDER, radius=0.08,
         line_w=1.0):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x),
                                Inches(y), Inches(w), Inches(h))
    no_shadow(sh)
    try:
        sh.adjustments[0] = radius
    except Exception:
        pass
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(line_w)
    return sh

def chip(slide, x, y, w, h, text, size=12, color=WHITE, fill=CARD,
         line=BORDER, bold=True, align=PP_ALIGN.CENTER, radius=0.25,
         sub=None, sub_size=9.5, sub_color=MUTED):
    c = card(slide, x, y, w, h, fill=fill, line=line, radius=radius)
    tf = c.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for m in ("margin_left", "margin_right"):
        setattr(tf, m, Inches(0.08))
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    if sub:
        p2 = tf.add_paragraph()
        p2.alignment = align
        p2.space_before = Pt(1)
        r2 = p2.add_run()
        r2.text = sub
        r2.font.name = FONT
        r2.font.size = Pt(sub_size)
        r2.font.color.rgb = sub_color
    return c

def conn(slide, x1, y1, x2, y2, color=MUTED, w=1.0, dash=None):
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1),
                                    Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = color
    cn.line.width = Pt(w)
    if dash:
        cn.line.dash_style = dash
    no_shadow(cn)
    return cn

def circle(slide, cx, cy, d, text=None, size=12, fill=None, line=BLUE,
           bold=True, color=WHITE):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - d / 2),
                                Inches(cy - d / 2), Inches(d), Inches(d))
    no_shadow(sh)
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1.25)
    if text is not None:
        tf = sh.text_frame
        tf.word_wrap = False
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, m, 0)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = text
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return sh

def header(slide, num_label, title, subtitle=None):
    w_ = min(4.6, 0.1 * len(num_label) + 0.5)
    chip(slide, 0.55, 0.32, w_, 0.34, num_label, size=10.5, color=BLUE,
         fill=None, line=BLUE_D, bold=True)
    tx(slide, 0.55, 0.78, 11.6, 0.75,
       [dict(text=title, size=28, color=WHITE, bold=True)])
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.58),
                                 Inches(1.42), Inches(0.85), Inches(0.055))
    no_shadow(bar)
    bar.fill.solid()
    bar.fill.fore_color.rgb = BLUE
    bar.line.fill.background()
    if subtitle:
        tx(slide, 1.55, 1.42, 10.6, 0.4,
           [dict(text=subtitle, size=13, color=MUTED, italic=True)],
           align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

def footer(slide, num):
    tx(slide, 0.55, 7.02, 6.0, 0.3,
       [dict(text="EKATVA AI  |  Many Apps. One Intelligence.", size=9,
             color=MUTED)])
    tx(slide, 12.0, 7.02, 0.8, 0.3,
       [dict(text="%02d / 15" % num, size=9, color=MUTED)],
       align=PP_ALIGN.RIGHT)

def arrow_right(slide, x, y, w, h, fill=BLUE):
    sh = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y),
                                Inches(w), Inches(h))
    no_shadow(sh)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    return sh

def logo_mark(slide, cx, cy, scale=1.0, hub_text="EKATVA"):
    """Minimal geometric unity mark: several nodes converging into one core."""
    d_node = 0.62 * scale
    d_hub = 1.55 * scale
    R = 1.55 * scale
    labels = ["Mail", "Cal", "Notes", "Chat", "Drive"]
    for i in range(5):
        a = math.radians(-90 + i * 72)
        nx, ny = cx + R * math.cos(a), cy + R * math.sin(a)
        conn(slide, cx + (d_hub / 2) * math.cos(a), cy + (d_hub / 2) * math.sin(a),
             nx - (d_node / 2) * math.cos(a), ny - (d_node / 2) * math.sin(a),
             color=BLUE_D, w=1.2)
        circle(slide, nx, ny, d_node, labels[i], size=8.5 * scale + 2,
               fill=CARD, line=MUTED, color=SOFT)
    hub = circle(slide, cx, cy, d_hub, None, fill=CARD_HI, line=BLUE)
    tf = hub.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = hub_text
    r.font.name = FONT
    r.font.size = Pt(15 * scale + 2)
    r.font.bold = True
    r.font.color.rgb = WHITE
    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = "ONE CORE"
    r2.font.name = FONT
    r2.font.size = Pt(7.5 * scale + 2)
    r2.font.color.rgb = BLUE
    p2.alignment = PP_ALIGN.CENTER
    return hub

def fade(slide, speed="slow"):
    """Subtle fade transition."""
    try:
        from lxml import etree
        sld = slide._element
        el = etree.SubElement(
            sld, "{http://schemas.openxmlformats.org/presentationml/2006/main}transition")
        el.set("spd", speed)
        etree.SubElement(el, "{http://schemas.openxmlformats.org/presentationml/2006/main}fade")
        # move transition after clrMapOvr (schema order)
        clr = sld.find("{http://schemas.openxmlformats.org/presentationml/2006/main}clrMapOvr")
        if clr is not None:
            clr.addnext(el)
    except Exception:
        pass

def notes(slide, text):
    tf = slide.notes_slide.notes_text_frame
    tf.text = text

def bullets(slide, x, y, w, h, items, size=13, color=SOFT, gap=6,
            marker="\u2022  ", marker_color=BLUE, line=None):
    paras = []
    for it in items:
        d = dict(text=marker + it, size=size, color=color,
                 space_after=gap)
        if line:
            d["line"] = line
        paras.append(d)
    tx(slide, x, y, w, h, paras)

# ═══════════════════════════ SPEAKER NOTES ═══════════════════════════

N1 = """PRESENTER: VIPUL KRISHNA PANDEY (Slides 1-8)

1) WHAT THIS SLIDE MEANS
This is our opening slide. We welcome the audience, ask one relatable question, and introduce EKATVA AI as a proposed startup idea for our entrepreneurship project — not a launched company.

2) SPEAKING SCRIPT
"Good morning everyone. Before we begin, let me ask you one question — what if you never had to open 10 different apps to get your daily work done? Think about your own day. Gmail for email. Google Calendar for classes. WhatsApp for friends. Notion for notes. We keep switching between apps all day long. Our startup idea, EKATVA AI, attacks exactly this pain point. The word 'Ekatva' means oneness or unity — many things becoming one. Our tagline says it all: 'Many Apps. One Intelligence.' EKATVA is a proposed AI-powered Personal Digital Agent — one interface where you give one command, and AI does the work across your connected apps. I am Vipul Krishna Pandey, enrollment BTECH/15202/26. My teammate is Mohan Sharma, enrollment BTECH/15200/26, and we are presenting this idea under the guidance of Dr. R.S. Pandey."

3) TECHNICAL TERMS SIMPLY
- AI Agent: a program that does not just answer questions; it can plan steps and take actions on your behalf, with your permission.
- Personal Digital Agent: an assistant that works inside your apps and does small jobs for you.

4) EXAMPLE
If you say "organise my tomorrow", instead of you opening Calendar, email and notes apps yourself, EKATVA would prepare that plan in one go.

5) TRANSITION
"In the next slide, we will look closely at the problem everyone faces — our digital life is fragmented."

6) POSSIBLE FACULTY QUESTIONS
- Why the name EKATVA? (Sanskrit word for oneness/unity — many apps, one intelligence.)
- Is this an existing product? (No, it is a proposed concept and MVP plan, presented for this project.)"""

N2 = """PRESENTER: VIPUL KRISHNA PANDEY (Slides 1-8)

1) WHAT THIS SLIDE MEANS
This slide shows the core pain point: our digital life is fragmented. We use too many apps that do not talk to each other smoothly, so we waste time managing apps instead of doing real work.

2) SPEAKING SCRIPT
"Let me quickly summarise the problem in six points. First, we use far too many applications every single day. Second, we constantly switch between them — studies call this 'context switching', and every switch costs us focus. Third, because everything is scattered, important emails and deadlines actually get missed. Fourth, we manually copy things around — we type the same task into our notes, our calendar and our reminders. Fifth, these apps were never really designed to work together smoothly. And sixth, the result: people spend their time managing the apps rather than completing the actual work. If you look at the visual — a person standing in the middle of disconnected apps — that is honestly a picture of all of us."

3) TECHNICAL TERMS SIMPLY
- Context switching: moving your attention from one app to another again and again. It sounds small, but studies show it drains concentration.

4) EXAMPLE
Think of a student: syllabus in one app, deadline in a second app, notes in a third. Before even starting the assignment, 15 minutes of setup is gone.

5) TRANSITION
"So the problem is clear. Now let me show you our solution — One Intelligence, Multiple Applications."

6) POSSIBLE FACULTY QUESTIONS
- Is this really a big problem, or just a first-world inconvenience? (For students and professionals juggling 8-10 apps daily, productivity loss is real and measurable in hours per week.)
- Have you quantified the problem? (We plan to validate this through surveys in our user research; we are not claiming any market-size numbers today.)"""

N3 = """PRESENTER: VIPUL KRISHNA PANDEY (Slides 1-8)

1) WHAT THIS SLIDE MEANS
This slide introduces the solution: EKATVA as a single command interface that coordinates several connected applications and produces one result. It also clarifies the difference between a normal chatbot and an agent.

2) SPEAKING SCRIPT
"Our solution: EKATVA is a proposed AI-powered Personal Digital Agent. The idea is simple — a single interface where you type or speak one natural command. Behind the scenes, EKATVA coordinates the relevant connected applications, and you get one final result. The flow in the visual shows it: one interface, one command, multiple connected apps, coordinated execution, one result. Now, one important distinction. A chatbot is like a knowledgeable friend — you ask, it answers. An AI agent is like an assistant — you ask, and it actually does the job for you, step by step, with your permission. A chatbot can tell you how to organise your calendar. An agent like EKATVA would actually prepare that calendar for you — after showing you and asking for approval where needed."

3) TECHNICAL TERMS SIMPLY
- Natural-language command: speaking or typing in normal English, like 'organise my tomorrow', instead of clicking buttons or setting up workflows.
- Coordinated execution: the agent splits your request and runs the steps across different apps in the right order.

4) EXAMPLE
"Email my team the meeting summary" — a chatbot would draft text and stop. EKATVA would summarise, show you the draft, and send it from your email after your approval.

5) TRANSITION
"So how does one command actually turn into multiple completed tasks? Next slide: How EKATVA Works."

6) POSSIBLE FACULTY QUESTIONS
- Isn't this just a chatbot with extra steps? (No — a chatbot only replies with text; an agent plans, uses app integrations and performs authorised actions.)
- Who sees my data when the agent does this? (Only the user's own connected accounts, under permissions the user grants — covered in the challenges slide.)"""

N4 = """PRESENTER: VIPUL KRISHNA PANDEY (Slides 1-8)

1) WHAT THIS SLIDE MEANS
This slide walks through the working of EKATVA in 9 steps, plus a real-world example command, so the workflow is easy to visualise.

2) SPEAKING SCRIPT
"Let me walk you through how EKATVA would work, in plain steps. One — the user opens EKATVA. Two — gives a command by text or voice. Three — the AI understands the intention behind the words. Four — it breaks the request into smaller tasks. Five — it selects the relevant connected applications. Six — it checks permissions and prepares the actions. Seven — and this is important — the user approves any sensitive action first. Eight — the agent executes the supported tasks. Nine — EKATVA verifies what happened and reports the results clearly. Now look at the example on the right. The user says: 'Check my important emails, identify pending tasks, add them to Notion and organise my calendar for tomorrow.' Notice the approval step — nothing irreversible happens without the user's confirmation."

3) TECHNICAL TERMS SIMPLY
- Intent: what the user actually wants, even if the English is loose. 'Organise my tomorrow' means plan my schedule.
- Task planning: splitting a big request into small ordered steps.

4) EXAMPLE
The example command would trigger roughly: read authorised emails, list important items, propose tasks, populate Notion, then build tomorrow's schedule around free slots.

5) TRANSITION
"Since it can plan and act like this, the natural question is — what exactly can EKATVA do? Next slide."

6) POSSIBLE FACULTY QUESTIONS
- What if step 2 fails, e.g. the AI misunderstands? (The agent shows its plan in preview; the user can correct or cancel before anything executes.)
- Voice as well as text? (Yes, the design supports both — voice is simply converted to the same command input.)"""

N5 = """PRESENTER: VIPUL KRISHNA PANDEY (Slides 1-8)

1) WHAT THIS SLIDE MEANS
This slide markets the vision in four everyday categories: personal life, communication, student life and professional work. It clearly states these are proposed capabilities.

2) SPEAKING SCRIPT
"We have grouped EKATVA's proposed capabilities into four simple categories. Personal life: organise schedules, set reminders, manage daily plans. Communication: summarise important emails, prepare email replies, and identify urgent messages first. Student life: track assignments, organise study schedules, create task lists and manage deadlines. Professional work: organise meetings, coordinate tasks, prepare daily summaries and automate repetitive workflows. One honest disclaimer, which we feel is important: all of these are proposed capabilities. They depend on which integrations we support and what permissions each user grants. We are not claiming a finished product — this is what we would build towards, starting with the MVP."

3) TECHNICAL TERMS SIMPLY
- Integration: an app officially connected to EKATVA through its approved interface (API), like Gmail being connected.
- Permissions: the specific rights a user gives, e.g. 'read my calendar' but not 'delete my emails'.

4) EXAMPLE
A student's morning command — 'summarise my urgent emails and show my deadlines this week' — would cover both the Communication and Student categories in one shot.

5) TRANSITION
"Behind these capabilities sits a technology stack. Don't worry — I will keep it simple. Next slide."

6) POSSIBLE FACULTY QUESTIONS
- Which of these is hardest to build? (Reliable multi-step planning is the hardest; single actions like creating a calendar event are more straightforward.)
- Will all apps be supported? (No — we start with official APIs like Gmail, Calendar and Notion, then expand.)"""

N6 = """PRESENTER: VIPUL KRISHNA PANDEY (Slides 1-8)

1) WHAT THIS SLIDE MEANS
This is our ONLY technology-stack slide. It shows the proposed stack in layers, explained in simple language so non-technical faculty can follow it.

2) SPEAKING SCRIPT
"Let me explain the proposed technology in one minute, layer by layer. At the top, the user-facing apps: React Native with Expo for the mobile app, and Next.js with React for the web app. Below that, the backend: Node.js with TypeScript and Fastify — this layer simply manages everything happening behind the scenes. Then the AI layer: OpenAI or Claude models to understand what the user means — and LangGraph, another technical term, which coordinates the agent's multi-step workflow, essentially the brain of the system that plans which step happens when. The data layer uses PostgreSQL with Prisma and a feature called pgvector that lets us store useful memory. The integrations layer uses OAuth 2.0 for secure authorised connections, standard REST APIs, and MCP — a common protocol for connecting tools. And background jobs and deployment: Redis with BullMQ handle tasks that run in the background, and Docker with AWS, Railway or Vercel handle deployment. That is the entire stack in one picture."

3) TECHNICAL TERMS SIMPLY
- Backend: the server-side part users never see, which does the actual work.
- OAuth 2.0: the industry-standard 'Login with...' system — the user explicitly authorises EKATVA, and sensitive passwords are never shared.
- pgvector: a database feature used to store and retrieve meaningful context.
- Background tasks: jobs like sending a summary that run without the user waiting.

4) EXAMPLE
The morning-email scenario flows through all layers: app command, backend receives it, AI model interprets, LangGraph plans, integrations execute, database records the history.

5) TRANSITION
"Technology aside, no startup exists in a vacuum — so let's honestly examine the existing market. Next slide."

6) POSSIBLE FACULTY QUESTIONS
- Why React Native and not Flutter? (React Native lets our team share knowledge with the React web side; the choice is proposed, not final.)
- Is AI expensive to run? (Yes — which is why our business model includes usage limits; discussed later.)
- Why LangGraph specifically? (It is a widely used orchestration framework for multi-step agent workflows with good community support.)"""

N7 = """PRESENTER: VIPUL KRISHNA PANDEY (Slides 1-8)

1) WHAT THIS SLIDE MEANS
An honest look at existing solutions — Zapier, n8n, ChatGPT, Lindy and Power Automate — what they already do well, and the user-experience opportunity EKATVA would target.

2) SPEAKING SCRIPT
"We researched the existing market honestly, because a good idea must survive real competitors. Zapier connects applications and automates workflows. n8n provides very flexible automation and AI-friendly workflows. ChatGPT already offers connected apps and supported actions. Lindy provides AI assistants for workflows, and Microsoft Power Automate automates business processes. We want to be clear: these products already deliver overlapping capabilities, and we respect them. So where is the space for EKATVA? Not in raw features, but in experience. We propose: a simpler everyday personal-assistant experience for normal users; a unified interface for common personal tasks; natural-language coordination across supported applications; clear visibility into exactly what the AI is doing; and a friendly permission and approval system. Many existing tools are powerful but geared towards builders and power users. EKATVA would target the everyday user — especially students and professionals."

3) TECHNICAL TERMS SIMPLY
- Power user: someone technical who enjoys configuring complex tools.
- Onboarding: the first-time experience of a new user who wants value in minutes, not hours.

4) EXAMPLE
Setting up a Zapier workflow involves triggers, actions and mapping fields. An everyday user may prefer simply writing: 'When I get an important email, make me a task.'

5) TRANSITION
"Based on that positioning, here is what we believe makes EKATVA different — and I want to present it fairly."

6) POSSIBLE FACULTY QUESTIONS
- Can't ChatGPT already do all this? (It supports connected apps and actions, yes. Our focus is the everyday, Indian, multi-app personal experience with a clear transparency-and-approval layer — an experience we want to build better, not a claim that others cannot do it.)
- Why did you not compare exact pricing? (Pricing changes frequently; we compared capability and experience categories instead.)"""

N8 = """PRESENTER: VIPUL KRISHNA PANDEY (Slides 1-8)

1) WHAT THIS SLIDE MEANS
Our proposed differentiation in seven points, plus a fair three-way comparison: a conventional chatbot, a workflow automation tool, and EKATVA.

2) SPEAKING SCRIPT
"Seven things we would build EKATVA around. One — a single command centre for everyday digital activities. Two — natural-language interaction, so users never have to manually build complex workflows. Three — coordination of tasks across connected apps. Four — personal context and preferences, always under user controls. Five — proactive reminders and daily briefings. Six — a transparent execution history, so users always see what the AI did. Seven — user approval before any sensitive action. On the right, a fair comparison. A conventional chatbot understands language but mostly stops at conversation — it does not take actions. A workflow automation tool does connect apps and take actions, but usually requires the user to build and maintain workflows manually. EKATVA is designed to combine both — understand the language and do the coordinated work — with transparency and approvals on top. Our line: EKATVA is not just about connecting apps. It is about making your digital life easier to manage."

3) TECHNICAL TERMS SIMPLY
- Proactive: the assistant reminds or suggests before you ask, e.g. a morning briefing.
- Execution history: a visible log of every action taken — what, when, in which app.

4) EXAMPLE
Ask 'send the report to my guide' — chatbot replies with advice; automation tool needs a pre-built rule; EKATVA would find the report, draft the email, show it and send after approval.

5) TRANSITION — HANDOVER
"This completes the product story. Now I would like to hand over to my teammate Mohan Sharma, who will take you through our target customers, MVP plan, business model and roadmap. Over to you, Mohan."

6) POSSIBLE FACULTY QUESTIONS
- Is the comparison fair to other products? (It groups typical products by category and notes that exact features vary; ChatGPT, for example, does support connected actions.)
- What stops a big player from copying you? (Execution speed, local relevance, focus on a specific everyday experience — and big players usually build broadly, not deep on one niche.)"""

N9 = """PRESENTER: MOHAN SHARMA (Slides 9-15)

1) WHAT THIS SLIDE MEANS
This slide defines who would actually use EKATVA — five segments — and marks our primary initial target: students and working professionals.

2) SPEAKING SCRIPT
"Thank you, Vipul. Good morning everyone. Let me now cover the business side, starting with customers. We see five natural segments. College students — they juggle assignments, deadlines and schedules, and want one place to manage it. Working professionals — they coordinate meetings, emails and daily tasks and hate app-hopping. Freelancers — they handle multiple clients, so communication and scheduling automation saves them real time. Entrepreneurs — they want one view of their business communication. And small business owners — who need simple automation without hiring technical staff. Our primary initial target is clear: students and working professionals who regularly use multiple digital applications. As a student team, we know this segment best — classmates around us are our first user-research pool, and professionals are close behind."

3) TECHNICAL TERMS SIMPLY
- Customer segment: a group of similar customers you design and market for.
- Initial target market: the segment you attack first before expanding.

4) EXAMPLE
A final-year student during placement season handles emails, interviews, deadlines and notes — an ideal first user for EKATVA.

5) TRANSITION
"Knowing the target, the next question is: what exactly do we build first? That brings us to the MVP."

6) POSSIBLE FACULTY QUESTIONS
- Why start with students? (We understand the segment, reach is cheap on campus, and students are heavy multi-app users and early adopters.)
- How large is each segment? (We will not quote market-size figures without proper research; validation via campus surveys is part of our plan.)
- Will professionals pay? (That is a hypothesis we would test with a free-first funnel and upgrade tracking.)"""

N10 = """PRESENTER: MOHAN SHARMA (Slides 9-15)

1) WHAT THIS SLIDE MEANS
This slide explains our realistic Minimum Viable Product — what would be in the first version — plus a practical demo scenario and its expected sequence.

2) SPEAKING SCRIPT
"MVP means Minimum Viable Product — the smallest useful version we would build first. Notice the wording on this slide: none of this is completed development. This is our proposed build plan. The MVP would include user registration and login; connecting Gmail, Google Calendar and Notion; a natural-language command interface; AI understanding of instructions; multi-step task planning; email summarisation; email reply preparation; calendar event creation; Notion task creation; a full execution history; and — again — approval before sensitive actions. Here is a practical demonstration scenario: 'Check my emails, find important tasks, add them to Notion and organise my calendar for tomorrow.' The expected sequence: read authorised emails, identify important tasks, prepare a task list, ask for approval, create the approved Notion tasks, review calendar slots, suggest or create an approved schedule, and finally display a summary. Thirteen features, one clear flow — ambitious but realistic with the stack Vipul showed."

3) TECHNICAL TERMS SIMPLY
- MVP: the smallest version that delivers the core value; you build it first, then improve with feedback.
- Execution history: a log showing exactly what the agent did and when — this builds trust.

4) EXAMPLE
If it reads three important emails — one assignment submission, one meeting invitation and one fee deadline — it would propose exactly three tasks and one calendar plan for approval.

5) TRANSITION
"A product needs a business engine. Next slide: how EKATVA would earn money."

6) POSSIBLE FACULTY QUestions
- Can three students build this? (Not all 13 features at once — we would sequence them, starting with the email-to-Notion-calendar flow as the core demo. The stack we chose is starter-friendly.)
- How long would the MVP take? (Rough estimate, assuming a usable alpha in 3-5 months of focused part-time work; refined after a technical spike.)
- Which feature is cut first if time is short? (Voice input — it is a layer over text commands, not the core value.)"""

N11 = """PRESENTER: MOHAN SHARMA (Slides 9-15)

1) WHAT THIS SLIDE MEANS
Our proposed freemium business model with four illustrative plans — Free, Pro, Premium and Business/Team — with honest disclaimers that the prices are assumptions to validate.

2) SPEAKING SCRIPT
"How would EKATVA earn money? A freemium subscription model — free to start, paid for power. The proposed Free plan would give limited monthly AI tasks, 2-3 connected apps and basic automation — enough to experience the value. The proposed Pro plan at 299 rupees per month would unlock more connected applications, higher task limits, personal preferences and memory, and recurring workflows. The Premium plan at 699 rupees per month would add advanced multi-step workflows, proactive assistance and priority execution, with higher usage limits. A Business/Team plan would offer team workflows, shared automation and administrative controls, priced on request. One important note: these prices are proposed assumptions for the business plan. They would be validated through customer research and our actual AI infrastructure costs — because AI calls cost real money per task, pricing must cover them sustainably."

3) TECHNICAL TERMS SIMPLY
- Freemium: free forever version plus paid premium plans — the standard modern SaaS model.
- Recurring workflows: automations that run on a schedule, like a tasks-digest every evening.
- SaaS: software as a subscription service.

4) EXAMPLE
Like Spotify's model: the free tier creates the habit; heavy users get more convenience by paying. Our equivalent of 'heavy' is more apps connected and more AI tasks per month.

5) TRANSITION
"Plans only matter if users arrive. Next slide: how we would attract our first customers."

6) POSSIBLE FACULTY QUESTIONS
- Is Rs 299 justified when ChatGPT has free tiers? (ChatGPT's free tier is broad but general; we would pitch deep personal-workflow convenience — and the price itself is an assumption to test, not a commitment.)
- What is a unit's real cost? (Primarily AI API tokens per task plus hosting; we would instrument usage per user from day one to price profitably.)
- What if users expect everything free? (Usage limits and the experience of a successful automation drive the upgrade decision — validated through A/B tests.)"""

N12 = """PRESENTER: MOHAN SHARMA (Slides 9-15)

1) WHAT THIS SLIDE MEANS
Our three-phase marketing plan, from campus launch to professionals to small businesses — plus the five-stage customer journey.

2) SPEAKING SCRIPT
"Our customer acquisition plan has three phases. Phase 1 is the campus launch: college campus promotion, student ambassadors, live demonstrations, Instagram and LinkedIn content, and a referral programme. Live demos are our strongest weapon — watching one command create tasks and a calendar is a 'wow' moment. Phase 2 moves to freelancers and professionals: presence in productivity communities, short demonstration videos, and user testimonials once we actually have users. Phase 3 reaches small businesses with team subscriptions and partnerships with relevant service providers. At the bottom is the journey every user walks: awareness, free signup, the first successful automation — that is the magic moment — then regular usage, and finally a paid subscription. Our entire funnel is designed to engineer that first 'it actually worked' moment as early as possible."

3) TECHNICAL TERMS SIMPLY
- Referral programme: existing users invite friends, both get benefits — cheap, trusted growth.
- Conversion funnel: the path from 'heard about it' to 'paying customer'.

4) EXAMPLE
Student ambassadors demo EKATVA in the canteen queue; a classmate signs up free on the spot, connects Gmail and Notion, and runs the email-to-tasks command the same evening.

5) TRANSITION
"Of course, no plan survives reality without problems. Next slide: the honest challenges."

6) POSSIBLE FACULTY QUESTIONS
- What is the customer acquisition cost? (Not yet measured; campus-phase costs are mostly time and events, which is why they suit a student startup.)
- Why not paid ads first? (Paid ads before product-market fit waste money; communities and demos give honest feedback for free.)"""

N13 = """PRESENTER: MOHAN SHARMA (Slides 9-15)

1) WHAT THIS SLIDE MEANS
Real challenges a startup like this faces — privacy, AI mistakes, app restrictions, AI costs and user trust — with concrete proposed solutions for each.

2) SPEAKING SCRIPT
"Let me present the uncomfortable truths. One, privacy and security. Solution: OAuth authentication — that means the user explicitly authorises us through the app provider, we store encrypted tokens, we request only limited permissions, and we keep audit logs of every access. Two, AI mistakes. AI is not perfect, so: we verify actions, show clear execution previews before anything runs, and require user approval for sensitive actions. Three, application restrictions — some apps limit external access. Solution: start strictly with official APIs and supported integrations, not hacks. Four, AI operating costs — running AI models costs money on every task. Solution: usage limits, efficient models, and subscription plans that cover costs. Five, user trust — people must believe an AI can touch their email. Solution: a transparent activity history and full control over connected accounts. One key point: connecting an email address or phone number does not automatically grant access to every application — each service needs its own separate authorised integration."

3) TECHNICAL TERMS SIMPLY
- OAuth: 'delegate limited access without sharing your password' — standard technology used by Google sign-in.
- Audit log: a private record of exactly who did what, when — for accountability.
- Sensitive action: anything hard to undo or private, like sending an email.

4) EXAMPLE
Before EKATVA sends any reply, the user sees: 'I will send this reply to Prof. Gupta — [draft] — approve or edit?' Nothing goes out silently.

5) TRANSITION
"And beyond the challenges, here is where we would take EKATVA in the long term — the roadmap."

6) POSSIBLE FACULTY QUESTIONS
- What if Gmail access is revoked by Google? (It is a real dependency risk; that is why official APIs, good standing and multi-provider choices matter.)
- Who can see user data? (Design principle: data stays with the user's own accounts and encrypted storage; no manual browsing, every access logged.)"""

N14 = """PRESENTER: MOHAN SHARMA (Slides 9-15)

1) WHAT THIS SLIDE MEANS
A realistic five-phase roadmap: from MVP, to more integrations, to a personal AI manager, to advanced automation, and finally a unified AI ecosystem.

2) SPEAKING SCRIPT
"Here is our roadmap, kept deliberately realistic. Phase 1 is the MVP you saw: Gmail, Google Calendar, Notion, and the basic AI agent. Phase 2, more integrations: expand supported applications and workflows based on what users ask for most. Phase 3, the personal AI manager: deepen user context, remember preferences, and offer proactive reminders — the assistant becomes personal. Phase 4, advanced automation: explore approved browser and device automation, only where it is technically and legally supported — we are aware there are legitimate limits here. Phase 5, the unified AI ecosystem: build toward a comprehensive personal digital agent — essentially, your entire digital world on one command. Notice the honest note on the timeline: phases 4 and 5 depend on external platforms and regulations. We are showing ambition, but we are not promising magic."

3) TECHNICAL TERMS SIMPLY
- Proactive reminders: the assistant tells you what matters before you ask, like a briefing.
- Device automation: letting the assistant perform approved steps inside apps or on the device — powerful, so it needs strict consent and legal care.

4) EXAMPLE
Phase 3 in real life: at 7 p.m. EKATVA reminds — 'tomorrow you have a submission and a seminar; shall I prepare both?'

5) TRANSITION
"And that brings us to the close: The Future Is Unified."

6) POSSIBLE FACULTY QUESTIONS
- Which phase has the biggest risk? (Phase 4 — device automation depends on platform rules; we have kept it exploratory on purpose.)
- What does 'unified ecosystem' mean concretely? (Growing integrations, memory and automation until one assistant reliably handles a user's whole day.)
- What is the timeline for Phase 1? (MVP alpha estimate: 3-5 months of focused work — an estimate, refined after the first technical prototype.)"""

N15 = """PRESENTER: MOHAN SHARMA (Slides 9-15)

1) WHAT THIS SLIDE MEANS
The closing slide: one memorable line, the brand, a thank-you, and the final thought that captures our motivation.

2) SPEAKING SCRIPT
"To conclude: today, we manage dozens of applications. Tomorrow, we should only need one intelligent assistant. That assistant is what EKATVA AI proposes to become — a Personal Digital Agent that brings Gmail, Calendar, Notion and more together behind one intelligent interface, where one command gets multiple authorised tasks done. Our final thought is one sentence: our goal is not to create another app that people have to manage. Our goal is to create an intelligent assistant that helps them manage everything else. Thank you very much — and we would both be happy to answer your questions."

3) TECHNICAL TERMS SIMPLY
- Authorised tasks: actions EKATVA performs only through integrations the user has explicitly approved, with sensitive steps confirmed by the user.

4) EXAMPLE
Recall the opening question — 10 apps a day. EKATVA's answer: one interface, one command, one intelligence.

5) CLOSING NOTE
Keep this slide on screen during Q&A; the tagline 'Many Apps. One Intelligence.' reinforces the brand until the session ends — Vipul joins for the joint Q&A.

6) POSSIBLE FACULTY QUESTIONS
- What would make you abandon this idea? (If user research shows neither students nor professionals care, or if authorised integrations are not technically workable — honest kill-criteria.)
- What is the single next step after today? (Start the MVP core prototype: Gmail to Notion task creation with approval, then validate on campus.)"""

NOTES = [N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11, N12, N13, N14, N15]

# ═══════════════════════════ SLIDE 1 — TITLE ═══════════════════════════
s = new_slide(); bg(s); fade(s)

# faint decorative rings
for d, col in [(5.6, RGBColor(0x1B, 0x2E, 0x6E)), (8.4, RGBColor(0x14, 0x22, 0x58))]:
    ring = circle(s, 2.55, 3.95, d, None, line=col)
    ring.line.width = Pt(1.1)

# opening question
tx(s, 1.1, 0.42, 11.1, 0.7, [dict(text="What if you never had to open 10 different apps"
                                       " to get your daily work done?",
                                  size=19, color=SOFT, italic=True, align=PP_ALIGN.CENTER)],
   anchor=MSO_ANCHOR.MIDDLE)

# brand block (right of hub)
tx(s, 4.85, 2.0, 7.9, 1.15, [dict(text="EKATVA AI", size=58, color=WHITE, bold=True)])
tx(s, 4.95, 3.18, 7.9, 0.55, [dict(text="Many Apps. One Intelligence.", size=21,
                                   color=BLUE, italic=True, bold=True)])
chip(s, 4.95, 3.92, 3.35, 0.42, "ENTREPRENEURSHIP PRESENTATION", size=11,
     color=PURPLE, fill=None, line=PURPLE_D, radius=0.5)

logo_mark(s, 2.55, 3.95, scale=1.0)

# presenter / faculty cards
card(s, 4.95, 4.78, 3.7, 1.85, fill=CARD, line=BORDER)
tx(s, 5.2, 4.95, 3.3, 1.6, [
    dict(text="PRESENTED BY", size=10, color=BLUE, bold=True, space_after=6),
    dict(text="Vipul Krishna Pandey", size=14, color=WHITE, bold=True, space_after=1),
    dict(text="BTECH/15202/26", size=10.5, color=MUTED, space_after=7),
    dict(text="Mohan Sharma", size=14, color=WHITE, bold=True, space_after=1),
    dict(text="BTECH/15200/26", size=10.5, color=MUTED)])
card(s, 8.85, 4.78, 3.7, 1.85, fill=CARD, line=BORDER)
tx(s, 9.1, 4.95, 3.3, 1.6, [
    dict(text="FACULTY", size=10, color=BLUE, bold=True, space_after=6),
    dict(text="Dr. R.S. Pandey", size=14, color=WHITE, bold=True, space_after=6),
    dict(text="Entrepreneurship", size=11, color=MUTED),
    dict(text="BTech Program", size=11, color=MUTED)])
tx(s, 4.95, 6.75, 7.6, 0.3, [dict(text="A proposed AI-powered Personal Digital Agent — startup concept, BTech entrepreneurship project",
                                  size=9.5, color=MUTED, italic=True)])
footer(s, 1)
notes(s, N1)

# ═══════════════════════════ SLIDE 2 — PROBLEM ═══════════════════════════
s = new_slide(); bg(s); fade(s)
header(s, "SLIDE 02  —  THE PROBLEM", "Our Digital Life Is Fragmented",
       "Everyday reality: too many apps, too much switching, not enough getting done")

card(s, 0.55, 1.75, 6.6, 4.95, fill=CARD, line=BORDER)
bullets(s, 0.95, 2.05, 5.9, 4.5, [
    "We use too many applications every day",
    "We repeatedly switch between apps — every switch costs focus",
    "Important emails and deadlines get missed",
    "We manually organise tasks, notes and schedules",
    "Different apps do not work together smoothly",
    "People spend time managing apps instead of completing actual work",
], size=14, gap=14)

# disconnected apps visual
cx, cy = 9.9, 4.2
apps = ["Gmail", "Calendar", "Notion", "WhatsApp", "Zoom", "Drive"]
pos = [(9.9, 2.35), (11.6, 3.3), (11.6, 5.1), (9.9, 6.05), (8.2, 5.1), (8.2, 3.3)]
for (x, y) in pos:
    conn(s, x, y, cx, cy, color=RGBColor(0x3A, 0x44, 0x6E), w=1.1,
         dash=MSO_LINE.DASH)
for (x, y), name in zip(pos, apps):
    circle(s, x, y, 1.05, name, size=10.5, fill=CARD, line=MUTED, color=SOFT)
circle(s, cx, cy, 1.35, "YOU", size=13, fill=CARD_HI, line=DANGER, color=WHITE)
tx(s, 8.15, 6.6, 3.5, 0.3, [dict(text="Apps work separately. You do the connecting.",
                                 size=10.5, color=DANGER, italic=True,
                                 align=PP_ALIGN.CENTER)])
footer(s, 2)
notes(s, N2)

# ═══════════════════════════ SLIDE 3 — SOLUTION ═══════════════════════════
s = new_slide(); bg(s); fade(s)
header(s, "SLIDE 03  —  OUR SOLUTION", "One Intelligence, Multiple Applications",
       "EKATVA — an AI-powered Personal Digital Agent with a one-command interface")

card(s, 0.55, 1.75, 6.35, 4.95, fill=CARD, line=BORDER)
tx(s, 0.95, 2.0, 5.6, 1.15, [dict(text="EKATVA is a proposed AI-powered Personal Digital Agent that provides ONE command interface to manage tasks across connected applications.",
                                  size=13.5, color=WHITE, line=1.15)])
bullets(s, 0.95, 3.3, 5.6, 3.2, [
    "One interface",
    "One natural-language command",
    "Multiple connected applications",
    "Coordinated task execution",
    "One final result",
], size=14, gap=10)

# vertical pipeline visual (right)
rx = 7.35
chip(s, rx, 1.75, 5.4, 0.55, "ONE NATURAL-LANGUAGE COMMAND", size=12.5,
     color=WHITE, fill=PURPLE_D, line=PURPLE)
arrow_right(s, rx + 2.45, 2.42, 0.5, 0.32, fill=PURPLE)
circle(s, rx + 2.7, 3.95, 1.85, "EKATVA", size=15, fill=CARD_HI, line=BLUE)
# mini label inside hub handled by circle()
apps3 = ["Gmail", "Calendar", "Notion"]
for i, name in enumerate(apps3):
    y = 3.15 + i * 0.85
    conn(s, rx + 3.55, 3.95, rx + 4.25, y + 0.28, color=BLUE_D, w=1.2)
    chip(s, rx + 4.25, y, 1.15, 0.56, name, size=10.5, fill=CARD, line=BORDER)
arrow_right(s, rx + 2.45, 5.05, 0.5, 0.32, fill=BLUE)
chip(s, rx, 5.5, 5.4, 1.2, "ONE  FINAL  RESULT", size=14, color=WHITE,
     fill=CARD, line=BLUE, sub="The task is done — organised, summarised or scheduled",
     sub_size=10.5, sub_color=SOFT)
footer(s, 3)
notes(s, N3)

# ═══════════════════════════ SLIDE 4 — HOW IT WORKS ═══════════════════════════
s = new_slide(); bg(s); fade(s)
header(s, "SLIDE 04  —  HOW EKATVA WORKS", "From One Command to Completed Tasks",
       "Nine steps from the user's words to verified results")

steps = [
    "User opens EKATVA",
    "Gives a text or voice command",
    "AI understands the intention",
    "Request divided into smaller tasks",
    "Agent selects relevant connected apps",
    "Permissions checked, actions prepared",
    "User approves sensitive actions",
    "Agent executes supported tasks",
    "EKATVA verifies and reports results",
]
gx, gy, cw, chh, gp = 0.55, 1.72, 2.06, 0.86, 0.16
for i, st in enumerate(steps):
    r_, c_ = divmod(i, 3)
    x = gx + c_ * (cw + gp); y = gy + r_ * (chh + gp)
    sc = card(s, x, y, cw, chh, fill=CARD, line=BORDER, radius=0.14)
    badge = circle(s, x + 0.34, y + chh / 2, 0.42, str(i + 1), size=11,
                   fill=CARD_HI, line=BLUE, color=WHITE)
    tx(s, x + 0.62, y + 0.08, cw - 0.72, chh - 0.16,
       [dict(text=st, size=10.5, color=SOFT, line=1.05)], anchor=MSO_ANCHOR.MIDDLE)

# example card (right)
ex = card(s, 7.3, 1.72, 5.48, 4.82, fill=CARD_HI, line=BLUE, radius=0.05)
tx(s, 7.55, 1.9, 5.0, 0.35, [dict(text="REAL-WORLD EXAMPLE", size=10.5,
                                  color=BLUE, bold=True)])
tx(s, 7.55, 2.28, 5.0, 1.05, [dict(text="\u201CCheck my important emails, identify pending tasks,"
                                        " add them to Notion and organise my calendar for tomorrow.\u201D",
                                   size=13, color=WHITE, italic=True, line=1.12)])
mini = [("AI understands the request", PURPLE_D),
        ("Tasks planned step by step", PURPLE_D),
        ("You preview and approve", BLUE),
        ("Agent executes and reports", BLUE)]
my = 3.5
for label, colr in mini:
    chip(s, 7.55, my, 4.42, 0.5, label, size=11.5, color=WHITE, fill=CARD,
         line=colr, align=PP_ALIGN.LEFT, radius=0.5)
    if my < 4.7:
        da = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(9.65),
                                Inches(my + 0.5), Inches(0.22), Inches(0.2))
        no_shadow(da)
        da.fill.solid()
        da.fill.fore_color.rgb = MUTED
        da.line.fill.background()
    my += 0.72
tx(s, 7.55, 6.35, 5.0, 0.3, [dict(text="Works the same with text or voice.", size=10,
                                  color=MUTED, italic=True)])
tx(s, 0.55, 6.7, 6.8, 0.3, [dict(text="Every step stays visible — sensitive actions always wait for user approval.",
                                 size=10.5, color=BLUE, italic=True)])
footer(s, 4)
notes(s, N4)

# ═══════════════════════════ SLIDE 5 — CAPABILITIES ═══════════════════════════
s = new_slide(); bg(s); fade(s)
header(s, "SLIDE 05  —  CAPABILITIES", "What Can EKATVA Do?",
       "Everyday capabilities grouped into four categories (proposed, subject to integrations and permissions)")

cats = [
    ("PERSONAL LIFE", ["Organise schedules", "Set reminders", "Manage daily plans"]),
    ("COMMUNICATION", ["Summarise important emails", "Prepare email replies",
                       "Identify urgent messages"]),
    ("STUDENT LIFE", ["Track assignments", "Organise study schedules",
                      "Create task lists", "Manage deadlines"]),
    ("PROFESSIONAL WORK", ["Organise meetings", "Coordinate tasks",
                           "Prepare daily summaries", "Automate repetitive workflows"]),
]
dims = [(0.55, 1.75), (6.85, 1.75), (0.55, 4.2), (6.85, 4.2)]
acolors = [PURPLE, BLUE_D, BLUE, PURPLE_D]
for (x, y), (nm, its), ac in zip(dims, cats, acolors):
    card(s, x, y, 5.95, 2.28, fill=CARD, line=BORDER)
    ic = circle(s, x + 0.5, y + 0.5, 0.52, None, fill=ac, line=None)
    tx(s, x + 0.95, y + 0.22, 4.7, 0.55, [dict(text=nm, size=15, color=WHITE,
                                               bold=True)], anchor=MSO_ANCHOR.MIDDLE)
    bullets(s, x + 0.5, y + 0.85, 5.2, 1.4, its, size=11.5, gap=4)

tx(s, 0.55, 6.65, 12.2, 0.35, [dict(text="Note: All capabilities are PROPOSED — they depend on supported integrations and the permissions each user grants.",
                                    size=11, color=GOLD, italic=True)])
footer(s, 5)
notes(s, N5)

# ═══════════════════════════ SLIDE 6 — TECH STACK ═══════════════════════════
s = new_slide(); bg(s); fade(s)
header(s, "SLIDE 06  —  TECHNOLOGY", "Proposed Technology Stack",
       "One stack, explained simply — the platform EKATVA would be built on")

bands = [
    ("USER APPS", [("React Native + Expo", "Mobile application"),
                   ("Next.js + React", "Web application")]),
    ("BACKEND", [("Node.js + TypeScript + Fastify", "Manages backend operations & core logic"),
                 ("", "")]),
    ("AI LAYER", [("OpenAI / Claude API", "Understand user instructions"),
                  ("LangGraph", "Coordinate multi-step agent workflows")]),
    ("DATA", [("PostgreSQL + Prisma + pgvector", "Store user & task data, agent memory"),
              ("", "")]),
    ("INTEGRATIONS", [("OAuth 2.0", "Secure authorised sign-in"),
                      ("REST APIs", "Standard app connections"),
                      ("MCP", "Common protocol for tool access")]),
    ("INFRASTRUCTURE", [("Redis + BullMQ", "Background task queue"),
                        ("Docker + AWS / Railway / Vercel", "Deployment")]),
]
by, bh, bgp = 1.68, 0.74, 0.095
for i, (nm, comps) in enumerate(bands):
    y = by + i * (bh + bgp)
    band = card(s, 0.55, y, 2.0, bh, fill=CARD_HI, line=BORDER, radius=0.18)
    tx(s, 0.62, y, 1.86, bh, [dict(text=nm, size=10, color=BLUE, bold=True,
                                   align=PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE)
    ncols = 3 if len(comps) == 3 else 2
    colw = (9.7 if ncols == 3 else 9.7)
    if len(comps) == 1 or (ncols == 2 and comps[1][0] == ""):
        chip(s, 2.72, y + 0.02, 9.7, bh - 0.04, comps[0][0], size=12.5,
             fill=CARD, line=BORDER, sub=comps[0][1])
    else:
        w_ = (9.7 - (ncols - 1) * 0.15) / ncols
        for j, (t_, d_) in enumerate(comps):
            chip(s, 2.72 + j * (w_ + 0.15), y + 0.02, w_, bh - 0.04, t_,
                 size=12.5, fill=CARD, line=BORDER, sub=d_, sub_size=9)
tx(s, 0.55, 6.72, 12.2, 0.3,
   [dict(text="Data flows top-down:  your command  \u2192  backend  \u2192  AI plan  \u2192  authorised app actions  \u2192  your result.",
         size=10.5, color=MUTED, italic=True)])
footer(s, 6)
notes(s, N6)

# ═══════════════════════════ SLIDE 7 — MARKET RESEARCH ═══════════════════════════
s = new_slide(); bg(s); fade(s)
header(s, "SLIDE 07  —  MARKET RESEARCH", "Existing Solutions & Our Opportunity",
       "Honest view of real products in the market today")

comps = [
    ("Zapier", "Connects applications and automates workflows"),
    ("n8n", "Flexible automation and AI-friendly workflows"),
    ("ChatGPT", "Connected apps and supported actions"),
    ("Lindy", "AI assistants for workflows"),
    ("Power Automate", "Automates business processes"),
]
cy0 = 1.75
for i, (nm, ds) in enumerate(comps):
    y = cy0 + i * 0.98
    card(s, 0.55, y, 6.5, 0.86, fill=CARD, line=BORDER, radius=0.16)
    chip(s, 0.75, y + 0.17, 1.85, 0.52, nm, size=12.5, fill=CARD_HI, line=BLUE_D)
    tx(s, 2.85, y, 4.05, 0.86,
       [dict(text=ds, size=11, color=SOFT, line=1.05)], anchor=MSO_ANCHOR.MIDDLE)

oc = card(s, 7.3, 1.75, 5.48, 4.92, fill=CARD_HI, line=BLUE, radius=0.05)
tx(s, 7.55, 1.92, 5.0, 0.4, [dict(text="THE EKATVA OPPORTUNITY", size=11,
                                  color=BLUE, bold=True)])
tx(s, 7.55, 2.32, 5.0, 0.65, [dict(text="These products already deliver overlapping capabilities."
                                        " Our opportunity is a better everyday experience:",
                                   size=10.5, color=SOFT, italic=True, line=1.1)])
bullets(s, 7.55, 3.08, 5.0, 3.5, [
    "Simpler everyday personal digital assistant",
    "One unified interface for common personal tasks",
    "Natural-language coordination across supported apps",
    "Clear visibility into what the AI is doing",
    "User-friendly permission & approval system",
], size=11.5, gap=9)
tx(s, 0.55, 6.75, 6.6, 0.3, [dict(text="Existing tools are powerful — but often built for technical users who design workflows.",
                                  size=10, color=MUTED, italic=True)])
footer(s, 7)
notes(s, N7)

# ═══════════════════════════ SLIDE 8 — DIFFERENTIATION ═══════════════════════════
s = new_slide(); bg(s); fade(s)
header(s, "SLIDE 08  —  WHY EKATVA?", "Our Differentiation",
       "What we would build EKATVA around — presented fairly")

pil = ["One command centre for everyday digital life",
       "Natural language instead of building workflows",
       "Coordination of tasks across connected apps",
       "Personal context with full user controls",
       "Proactive reminders & daily briefings",
       "Transparent execution history",
       "User approval before sensitive actions"]
for i, p_ in enumerate(pil):
    tx(s, 0.55, 1.68 + i * 0.6, 4.6, 0.55,
       [dict(text="\u2022  " + p_, size=11.5, color=SOFT)], anchor=MSO_ANCHOR.MIDDLE)

# comparison card
cc = card(s, 5.35, 1.68, 7.43, 4.45, fill=CARD, line=BORDER)
cols = [("Simple\nChatbot", MUTED), ("Automation\nTool", MUTED), ("EKATVA\n(proposed)", BLUE)]
colx = [8.25, 9.8, 11.3]
rowsc = [
    ("Understands natural language", "Yes", "Limited", "Yes"),
    ("Connects to your apps", "No", "Yes", "Planned"),
    ("Takes authorised actions", "No", "Yes", "Planned"),
    ("Shows execution history", "No", "Limited", "Planned"),
    ("Personal proactive assistant", "No", "No", "Planned"),
]
for x, (cnm, cc_) in zip(colx, cols):
    chip(s, x, 1.86, 1.45, 0.52, cnm.replace("\n", "  "), size=9.5,
         fill=CARD_HI, line=cc_ if cc_ == BLUE else BORDER, radius=0.5,
         align=PP_ALIGN.CENTER)
ry0 = 2.5
for i, (lbl, a, b, c_) in enumerate(rowsc):
    y = ry0 + i * 0.6
    if i:
        conn(s, 5.75, y, 12.5, y, color=RGBColor(0x22, 0x2E, 0x5A), w=0.75)
    tx(s, 5.55, y, 2.55, 0.56, [dict(text=lbl, size=10, color=SOFT, line=1.0)],
       anchor=MSO_ANCHOR.MIDDLE)
    for x, v in zip(colx, (a, b, c_)):
        good = v in ("Yes", "Planned")
        lim = v == "Limited"
        colr = BLUE if v == "Planned" else (WHITE if good else MUTED)
        tx(s, x - 0.05, y, 1.55, 0.56,
           [dict(text=("\u2713 " + v) if good else ("\u2013 " + v),
                 size=10, color=colr, bold=(not lim), italic=lim,
                 align=PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE)
tx(s, 5.35, 5.55, 7.43, 0.55,
   [dict(text="Typical products grouped by category — exact features vary by product. EKATVA entries are our proposed design.",
         size=8.5, color=MUTED, italic=True, line=1.05)])

qb = card(s, 0.55, 6.3, 12.23, 0.62, fill=PURPLE_D, line=PURPLE, radius=0.3)
tx(s, 0.8, 6.3, 11.7, 0.62,
   [dict(text="\u201CEKATVA is not just about connecting apps. It is about making your digital life easier to manage.\u201D",
         size=13.5, color=WHITE, bold=True, italic=True, align=PP_ALIGN.CENTER)],
   anchor=MSO_ANCHOR.MIDDLE)
footer(s, 8)
notes(s, N8)

# ═══════════════════════════ SLIDE 9 — TARGET CUSTOMERS ═══════════════════════════
s = new_slide(); bg(s); fade(s)
header(s, "SLIDE 09  —  TARGET CUSTOMERS", "Who Would Use EKATVA?",
       "Five initial customer segments — two primary")

seg = [
    ("College Students", "Assignments, deadlines & schedules — all in one place", True),
    ("Working Professionals", "Daily coordination without constant app-hopping", True),
    ("Freelancers", "Match client communication & schedules in one flow", False),
    ("Entrepreneurs", "One view across scattered business communication", False),
    ("Small Business Owners", "Simple automation without technical staff", False),
]
xw, gap = 2.32, 0.16
x0 = 0.55
for i, (nm, ds, prim) in enumerate(seg):
    x = x0 + i * (xw + gap)
    fillc = CARD_HI if prim else CARD
    lc = BLUE if prim else BORDER
    card(s, x, 1.8, xw, 2.9, fill=fillc, line=lc)
    ic = circle(s, x + xw / 2, 2.35, 0.62, None, fill=PURPLE_D if prim else CARD, line=PURPLE)
    tx(s, x + 0.12, 2.75, xw - 0.24, 0.75,
       [dict(text=nm, size=13, color=WHITE, bold=True, align=PP_ALIGN.CENTER, line=1.02)])
    tx(s, x + 0.15, 3.5, xw - 0.3, 1.1,
       [dict(text=ds, size=10, color=SOFT, align=PP_ALIGN.CENTER, line=1.1)])
    if prim:
        chip(s, x + 0.26, 4.32, xw - 0.52, 0.3, "PRIMARY", size=8.5,
             color=WHITE, fill=BLUE, line=None, radius=0.5)

cb = card(s, 0.55, 5.05, 12.23, 1.55, fill=CARD_HI, line=BLUE)
tx(s, 0.85, 5.22, 11.6, 1.25, [
    dict(text="PRIMARY INITIAL TARGET MARKET", size=11, color=BLUE, bold=True, space_after=4),
    dict(text="Students and working professionals who regularly use multiple digital applications.",
         size=15, color=WHITE, bold=True, space_after=4),
    dict(text="Both segments feel the fragmentation problem daily, adopt new tools quickly, and are reached efficiently through campuses and professional networks.",
         size=10.5, color=SOFT)])
footer(s, 9)
notes(s, N9)

# ═══════════════════════════ SLIDE 10 — MVP ═══════════════════════════
s = new_slide(); bg(s); fade(s)
header(s, "SLIDE 10  —  MVP", "Minimum Viable Product — Proposed Build Plan",
       "The first realistic version we would build (proposed — no features are completed yet)")

mvp = ["User registration & login", "Connect Gmail", "Connect Google Calendar",
       "Connect Notion", "Natural-language command interface",
       "AI understanding of instructions", "Multi-step task planning",
       "Email summarisation", "Email reply preparation",
       "Calendar event creation", "Notion task creation",
       "Execution history", "Approval before sensitive actions"]
gx0, gy0, cw2, chh2 = 0.55, 1.72, 2.94, 0.56
for i, m in enumerate(mvp):
    r_, c_ = divmod(i, 2)
    x = gx0 + c_ * (cw2 + 0.16); y = gy0 + r_ * (chh2 + 0.115)
    chip(s, x, y, cw2, chh2, "\u2713  " + m, size=10.5, color=WHITE, fill=CARD,
         line=BORDER, align=PP_ALIGN.LEFT, radius=0.3)
tx(s, 0.55, 6.72, 6.2, 0.35, [dict(text="13 proposed features — nothing shown is completed development.",
                                   size=10, color=GOLD, italic=True)])

# demo card
dc = card(s, 6.85, 1.72, 5.93, 4.92, fill=CARD_HI, line=BLUE, radius=0.05)
tx(s, 7.1, 1.88, 5.4, 0.35, [dict(text="DEMONSTRATION SCENARIO", size=11,
                                  color=BLUE, bold=True)])
tx(s, 7.1, 2.24, 5.4, 0.85, [dict(text="\u201CCheck my emails, find important tasks, add them"
                                       " to Notion and organise my calendar for tomorrow.\u201D",
                                  size=12.5, color=WHITE, italic=True, line=1.1)])
seq = ["Read authorised emails", "Identify important tasks", "Prepare a task list",
       "Ask for approval", "Create approved Notion tasks",
       "Review available calendar slots", "Suggest / create approved schedule",
       "Display final summary"]
sy = 3.05
for i, stp in enumerate(seq):
    chip(s, 7.1, sy, 5.42, 0.38, "%d.  %s" % (i + 1, stp), size=9.5, color=WHITE,
         fill=CARD, line=BORDER if i not in (3, 7) else BLUE,
         align=PP_ALIGN.LEFT, radius=0.5)
    sy += 0.435
footer(s, 10)
notes(s, N10)

# ═══════════════════════════ SLIDE 11 — BUSINESS MODEL ═══════════════════════════
s = new_slide(); bg(s); fade(s)
header(s, "SLIDE 11  —  BUSINESS MODEL", "How Will EKATVA Earn Money?",
       "Proposed freemium subscription model — illustrative pricing, to be validated")

plans = [
    ("FREE", "\u20B90", ["Limited monthly AI tasks", "2\u20133 connected apps",
                          "Basic automation"], False),
    ("PRO", "\u20B9299 /month", ["More connected applications", "Higher task limits",
                                  "Personal preferences & memory", "Recurring workflows"], True),
    ("PREMIUM", "\u20B9699 /month", ["Higher usage limits", "Advanced multi-step workflows",
                                     "Proactive assistance", "Priority execution"], False),
    ("BUSINESS / TEAM", "Custom", ["Team workflows", "Shared automation",
                                   "Administrative controls"], False),
]
px0, pw, pgap = 0.55, 2.98, 0.12
for i, (nm, pr, its, hi) in enumerate(plans):
    x = px0 + i * (pw + pgap)
    card(s, x, 1.8, pw, 4.4, fill=CARD_HI if hi else CARD,
         line=BLUE if hi else BORDER, line_w=1.5 if hi else 1.0)
    tx(s, x + 0.15, 2.0, pw - 0.3, 0.4, [dict(text=nm, size=12.5,
        color=BLUE if hi else SOFT, bold=True, align=PP_ALIGN.CENTER)])
    tx(s, x + 0.15, 2.42, pw - 0.3, 0.5, [dict(text=pr, size=17, color=WHITE,
        bold=True, align=PP_ALIGN.CENTER)])
    conn(s, x + 0.45, 3.05, x + pw - 0.45, 3.05, color=BORDER, w=0.75)
    bullets(s, x + 0.28, 3.2, pw - 0.5, 2.9, its, size=10.5, gap=7, line=1.05)
    if hi:
        chip(s, x + pw / 2 - 0.7, 5.55, 1.4, 0.34, "FLAGSHIP", size=9,
             color=NAVY, fill=BLUE, line=None, radius=0.5)

tx(s, 0.55, 6.45, 12.2, 0.55,
   [dict(text="Pricing note: Prices above are PROPOSED ASSUMPTIONS for the business plan — not established market prices. They would be validated through customer research and actual AI infrastructure costs.",
         size=10.5, color=GOLD, italic=True, line=1.1)])
footer(s, 11)
notes(s, N11)

# ═══════════════════════════ SLIDE 12 — MARKETING ═══════════════════════════
s = new_slide(); bg(s); fade(s)
header(s, "SLIDE 12  —  MARKETING", "How We Would Attract Users",
       "Three phases — campus to professionals to small businesses")

phases = [
    ("PHASE 1\nCAMPUS LAUNCH", ["College campus promotion", "Student ambassadors",
                                "Live demonstrations", "Instagram & LinkedIn",
                                "Referral programme"]),
    ("PHASE 2\nPROFESSIONALS", ["Freelancers & professionals", "Productivity communities",
                                "Short demonstration videos",
                                "User testimonials (once available)"]),
    ("PHASE 3\nSMALL BUSINESS", ["Team subscriptions",
                                 "Partnerships with relevant service providers"]),
]
px0, pw2, pgap2 = 0.55, 3.9, 0.24
for i, (nm, its) in enumerate(phases):
    x = px0 + i * (pw2 + pgap2)
    card(s, x, 1.72, pw2, 3.35, fill=CARD, line=PURPLE_D if i == 0 else BORDER)
    tx(s, x + 0.2, 1.85, pw2 - 0.4, 0.62,
       [dict(text=nm, size=12, color=BLUE, bold=True, line=1.05)])
    bullets(s, x + 0.2, 2.55, pw2 - 0.4, 2.4, its, size=10.5, gap=6)
    if i < 2:
        arrow_right(s, x + pw2 + 0.015, 3.15, 0.21, 0.3, fill=PURPLE)

tx(s, 0.55, 5.32, 6.0, 0.35, [dict(text="THE CUSTOMER JOURNEY", size=11,
                                   color=BLUE, bold=True)])
journey = ["Awareness", "Free Signup", "First Successful Automation",
           "Regular Usage", "Paid Subscription"]
jx = 0.55
widths = [1.55, 1.75, 3.05, 2.1, 2.7]
for i, (w_, jn) in enumerate(zip(widths, journey)):
    hi = (i == 2)
    chip(s, jx, 5.75, w_, 0.62, jn, size=11, color=NAVY if hi else WHITE,
         fill=BLUE if hi else CARD, line=None if hi else BORDER, radius=0.4)
    if i < 4:
        arrow_right(s, jx + w_ + 0.03, 5.93, 0.28, 0.26, fill=PURPLE)
    jx += w_ + 0.34
tx(s, 0.55, 6.6, 12.0, 0.35,
   [dict(text="The magic moment: the first time a user's command actually completes across their apps.",
         size=10.5, color=MUTED, italic=True)])
footer(s, 12)
notes(s, N12)

# ═══════════════════════════ SLIDE 13 — CHALLENGES & SOLUTIONS ═══════════════════════════
s = new_slide(); bg(s); fade(s)
header(s, "SLIDE 13  —  CHALLENGES", "Challenges & Solutions",
       "Realistic startup risks — and how we would address them")

ch = [
    ("Privacy & Security", "OAuth authentication, encrypted tokens, limited permissions, audit logs"),
    ("AI Mistakes", "Action verification, clear execution previews, user approvals"),
    ("Application Restrictions", "Start with official APIs and supported integrations only"),
    ("AI Operating Costs", "Usage limits, efficient models, subscription plans cover compute"),
    ("User Trust", "Transparent activity history & full control over connected accounts"),
]
y0, rh = 1.72, 0.82
for i, (c_, sl) in enumerate(ch):
    y = y0 + i * (rh + 0.12)
    card(s, 0.55, y, 4.3, rh, fill=CARD, line=DANGER, radius=0.12)
    tx(s, 0.8, y, 3.9, rh, [dict(text=c_, size=12.5, color=DANGER, bold=True)],
       anchor=MSO_ANCHOR.MIDDLE)
    arrow_right(s, 4.98, y + rh / 2 - 0.13, 0.5, 0.26, fill=PURPLE)
    card(s, 5.6, y, 7.18, rh, fill=CARD_HI, line=BORDER, radius=0.12)
    tx(s, 5.85, y, 6.7, rh, [dict(text=sl, size=12, color=WHITE, line=1.08)],
       anchor=MSO_ANCHOR.MIDDLE)
nb = card(s, 0.55, 6.42, 12.23, 0.5, fill=CARD, line=GOLD, radius=0.2)
tx(s, 0.8, 6.42, 11.7, 0.5,
   [dict(text="Key point: connecting an email address or phone number does NOT automatically grant access to every application — each service needs its own authorised integration and permissions.",
         size=9.5, color=GOLD, italic=True, align=PP_ALIGN.CENTER, line=1.05)],
   anchor=MSO_ANCHOR.MIDDLE)
footer(s, 13)
notes(s, N13)

# ═══════════════════════════ SLIDE 14 — ROADMAP ═══════════════════════════
s = new_slide(); bg(s); fade(s)
header(s, "SLIDE 14  —  ROADMAP", "The Road Ahead",
       "Five realistic phases — from MVP to a unified AI ecosystem")

phases5 = [
    ("PHASE 1\nMVP", "Gmail, Google Calendar,\nNotion + basic AI agent", False),
    ("PHASE 2\nMORE INTEGRATIONS", "Expand supported\napps & workflows", False),
    ("PHASE 3\nPERSONAL AI MANAGER", "User context, preferences,\nproactive reminders", False),
    ("PHASE 4\nADVANCED AUTOMATION", "Approved browser / device\nautomation where supported", False),    ("PHASE 5\nUNIFIED AI ECOSYSTEM", "Toward a comprehensive\npersonal digital agent", True),
]
tl = arrow_right(s, 0.55, 3.52, 12.23, 0.4, fill=PURPLE_D)
px0, pw3, pgap3 = 0.55, 2.32, 0.17
for i, (nm, ds, fin) in enumerate(phases5):
    x = px0 + i * (pw3 + pgap3)
    node = circle(s, x + pw3 / 2, 3.72, 0.4, str(i + 1), size=11,
                  fill=NAVY, line=BLUE, color=WHITE)
    chip(s, x, 1.95, pw3, 1.1, nm.replace("\n", "  —  "), size=9.5,
         color=WHITE, fill=CARD_HI if fin else CARD,
         line=BLUE if fin else BORDER, radius=0.15)
    cd = card(s, x, 4.15, pw3, 1.5, fill=CARD if not fin else CARD_HI,
              line=BLUE if fin else BORDER)
    for k, ln_ in enumerate(ds.split("\n")):
        tx(s, x + 0.15, 4.32 + k * 0.42, pw3 - 0.3, 0.4,
           [dict(text=ln_, size=10, color=SOFT, align=PP_ALIGN.CENTER)])
    conn(s, x + pw3 / 2, 3.05, x + pw3 / 2, 3.52, color=BORDER, w=1.0)
    conn(s, x + pw3 / 2, 3.92, x + pw3 / 2, 4.15, color=BORDER, w=1.0)
tx(s, 0.55, 5.95, 12.2, 0.75,
   [dict(text="Honest note: Phases 4\u20135 depend on external platforms, permissions and regulations — they are exploratory, not commitments.",
         size=11, color=GOLD, italic=True, align=PP_ALIGN.CENTER)])
footer(s, 14)
notes(s, N14)

# ═══════════════════════════ SLIDE 15 — CONCLUSION ═══════════════════════════
s = new_slide(); bg(s); fade(s)
for d, col in [(6.2, RGBColor(0x1B, 0x2E, 0x6E)), (9.0, RGBColor(0x14, 0x22, 0x58)),
               (11.8, RGBColor(0x10, 0x1C, 0x4A))]:
    ring = circle(s, SLIDE_W / 2, 3.1, d, None, line=col)
    ring.line.width = Pt(1.1)

tx(s, 1.2, 0.55, 10.93, 1.15,
   [dict(text="\u201CToday, we manage dozens of applications.",
         size=24, color=WHITE, bold=True, italic=True, align=PP_ALIGN.CENTER),
    dict(text="Tomorrow, we should only need one intelligent assistant.\u201D",
         size=24, color=WHITE, bold=True, italic=True, align=PP_ALIGN.CENTER)])

logo_mark(s, SLIDE_W / 2, 3.35, scale=0.75)

tx(s, 1.2, 4.85, 10.93, 0.6,
   [dict(text="EKATVA AI", size=36, color=WHITE, bold=True, align=PP_ALIGN.CENTER)])
tx(s, 1.2, 5.45, 10.93, 0.45,
   [dict(text="Many Apps. One Intelligence.", size=16, color=BLUE, italic=True,
         bold=True, align=PP_ALIGN.CENTER)])
tx(s, 1.2, 5.95, 10.93, 0.5,
   [dict(text="THANK YOU", size=20, color=WHITE, bold=True, align=PP_ALIGN.CENTER)])

ft = card(s, 2.15, 6.5, 9.0, 0.62, fill=PURPLE_D, line=PURPLE, radius=0.3)
tx(s, 2.4, 6.5, 8.5, 0.62,
   [dict(text="\u201COur goal is not to create another app that people have to manage.\u2014"
              "but an intelligent assistant that helps them manage everything else.\u201D",
         size=11.5, color=WHITE, italic=True, align=PP_ALIGN.CENTER, line=1.05)],
   anchor=MSO_ANCHOR.MIDDLE)
notes(s, N15)

# ─────────────────────────── SAVE & VERIFY ───────────────────────────
prs.core_properties.title = "EKATVA AI - Entrepreneurship Presentation"
prs.core_properties.author = "Vipul Krishna Pandey & Mohan Sharma"
prs.core_properties.subject = "Many Apps. One Intelligence."
out = r"D:\EKATVA\EKATVA_Entrepreneurship_Presentation.pptx"
prs.save(out)

# quick verification
chk = Presentation(out)
n_slides = len(chk.slides)
n_notes = sum(1 for sl in chk.slides if sl.has_notes_slide and sl.notes_slide.notes_text_frame.text.strip())
print("Saved:", out)
print("Slides:", n_slides)
print("Slides with notes:", n_notes)
words = 0
for sl in chk.slides:
    for sh in sl.shapes:
        if sh.has_text_frame:
            words += len(sh.text_frame.text.split())
print("Total on-slide words:", words)
