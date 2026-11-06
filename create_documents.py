"""
EKATVA AI — Document Generator
Creates two professional Word documents:
1. Presentation Speaking Script
2. Faculty Q&A Preparation Guide
"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os

# ── Color constants ──
DARK_NAVY = RGBColor(10, 25, 65)
BLACK = RGBColor(0, 0, 0)
ACCENT_BLUE = RGBColor(30, 80, 160)
LIGHT_GRAY_BG = "D9E2F3"
PRESENTER_BG = "FFF2CC"
HANDOVER_BG = "E2EFDA"
WHITE = RGBColor(255, 255, 255)

FONT_NAME = "Calibri"


def set_cell_shading(cell, color_hex):
    """Set background shading for a table cell."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_paragraph_font(paragraph, size=12, color=BLACK, bold=False, italic=False, font_name=FONT_NAME):
    """Apply font formatting to all runs in a paragraph."""
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.bold = bold
        run.font.italic = italic
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)


def add_styled_paragraph(doc, text, size=12, color=BLACK, bold=False, italic=False,
                         alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=6,
                         font_name=FONT_NAME):
    """Add a paragraph with complete styling."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    p.alignment = alignment
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = Pt(size * 1.5)
    return p


def add_mixed_paragraph(doc, parts, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                        space_before=0, space_after=6, line_spacing_pt=18):
    """Add a paragraph with mixed formatting (list of dicts with text, bold, color, size, italic)."""
    p = doc.add_paragraph()
    for part in parts:
        run = p.add_run(part.get("text", ""))
        run.font.name = part.get("font_name", FONT_NAME)
        run.font.size = Pt(part.get("size", 12))
        run.font.color.rgb = part.get("color", BLACK)
        run.font.bold = part.get("bold", False)
        run.font.italic = part.get("italic", False)
        run._element.rPr.rFonts.set(qn('w:eastAsia'), part.get("font_name", FONT_NAME))
    p.alignment = alignment
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = Pt(line_spacing_pt)
    return p


def style_heading(heading, level=1, color=DARK_NAVY, size=None, font_name=FONT_NAME):
    """Style a heading with custom font and color."""
    if size is None:
        size = {1: 18, 2: 15, 3: 13}.get(level, 12)
    for run in heading.runs:
        run.font.name = font_name
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.bold = True
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    heading.paragraph_format.space_before = Pt(12)
    heading.paragraph_format.space_after = Pt(6)
    return heading


def add_presenter_box(doc, presenter_name, bg_color=PRESENTER_BG):
    """Add a highlighted presenter assignment box."""
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    cell = table.cell(0, 0)
    set_cell_shading(cell, bg_color)
    p = cell.paragraphs[0]
    run = p.add_run(f"🎤 PRESENTER: {presenter_name}")
    run.font.name = FONT_NAME
    run.font.size = Pt(11)
    run.font.color.rgb = DARK_NAVY
    run.font.bold = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    # Remove table borders for a cleaner look - set to single thin border
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '  <w:top w:val="single" w:sz="4" w:space="0" w:color="BDB76B"/>'
        '  <w:left w:val="single" w:sz="4" w:space="0" w:color="BDB76B"/>'
        '  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="BDB76B"/>'
        '  <w:right w:val="single" w:sz="4" w:space="0" w:color="BDB76B"/>'
        '</w:tblBorders>'
    )
    tblPr.append(borders)
    doc.add_paragraph()  # spacing


def add_handover_box(doc):
    """Add a handover transition box."""
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    cell = table.cell(0, 0)
    set_cell_shading(cell, HANDOVER_BG)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run("🔄 HANDOVER TRANSITION")
    run.font.name = FONT_NAME
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(33, 100, 33)
    run.font.bold = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)

    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run("Vipul Krishna Pandey hands over to Mohan Sharma")
    run2.font.name = FONT_NAME
    run2.font.size = Pt(11)
    run2.font.color.rgb = RGBColor(33, 100, 33)
    run2.font.italic = True
    run2._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)

    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '  <w:top w:val="single" w:sz="6" w:space="0" w:color="548235"/>'
        '  <w:left w:val="single" w:sz="6" w:space="0" w:color="548235"/>'
        '  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="548235"/>'
        '  <w:right w:val="single" w:sz="6" w:space="0" w:color="548235"/>'
        '</w:tblBorders>'
    )
    tblPr.append(borders)
    doc.add_paragraph()


def add_speech_text(doc, text):
    """Add speech/script text with proper formatting - italic, indented."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.right_indent = Inches(0.3)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = Pt(20)

    run = p.add_run(f'"{text}"')
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(30, 30, 30)
    run.font.italic = False
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
    return p


def set_document_margins(doc, margin_inches=1.0):
    """Set margins for all sections."""
    for section in doc.sections:
        section.top_margin = Inches(margin_inches)
        section.bottom_margin = Inches(margin_inches)
        section.left_margin = Inches(margin_inches)
        section.right_margin = Inches(margin_inches)


def add_horizontal_line(doc):
    """Add a horizontal line separator."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        '  <w:bottom w:val="single" w:sz="6" w:space="1" w:color="CCCCCC"/>'
        '</w:pBdr>'
    )
    pPr.append(pBdr)


# ════════════════════════════════════════════════════════════════════
# DOCUMENT 1: PRESENTATION SPEAKING SCRIPT
# ════════════════════════════════════════════════════════════════════

def create_speaking_script():
    print("Creating Document 1: Presentation Speaking Script...")
    doc = Document()
    set_document_margins(doc)

    # ── Style defaults ──
    style = doc.styles['Normal']
    font = style.font
    font.name = FONT_NAME
    font.size = Pt(12)
    font.color.rgb = BLACK
    style.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)

    for level in range(1, 4):
        hs = doc.styles[f'Heading {level}']
        hs.font.name = FONT_NAME
        hs.font.color.rgb = DARK_NAVY
        hs.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)

    # ── TITLE PAGE ──
    for _ in range(6):
        doc.add_paragraph()

    title_p = add_styled_paragraph(doc, "EKATVA AI", size=36, color=DARK_NAVY, bold=True,
                                    alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_styled_paragraph(doc, "Presentation Speaking Script", size=22, color=ACCENT_BLUE, bold=False,
                          alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)

    add_horizontal_line(doc)

    add_styled_paragraph(doc, "Entrepreneurship Presentation | BTech", size=14, color=DARK_NAVY,
                          alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=16, space_after=8)

    add_styled_paragraph(doc, "Many Apps. One Intelligence.", size=16, color=ACCENT_BLUE, bold=True,
                          italic=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=20)

    add_horizontal_line(doc)

    # Presenter info
    add_styled_paragraph(doc, "", size=12, space_after=20)

    # Presenter assignment table
    table = doc.add_table(rows=3, cols=2)
    table.style = 'Table Grid'

    # Header row
    for i, text in enumerate(["Section", "Presenter"]):
        cell = table.cell(0, i)
        set_cell_shading(cell, "0A1941")
        p = cell.paragraphs[0]
        run = p.add_run(text)
        run.font.name = FONT_NAME
        run.font.size = Pt(11)
        run.font.color.rgb = WHITE
        run.font.bold = True
        run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    data_rows = [
        ("Slides 1 – 8", "Vipul Krishna Pandey"),
        ("Slides 9 – 15", "Mohan Sharma"),
    ]
    for r, (sec, pres) in enumerate(data_rows, 1):
        for c, text in enumerate([sec, pres]):
            cell = table.cell(r, c)
            if r % 2 == 0:
                set_cell_shading(cell, LIGHT_GRAY_BG)
            p = cell.paragraphs[0]
            run = p.add_run(text)
            run.font.name = FONT_NAME
            run.font.size = Pt(11)
            run.font.color.rgb = BLACK
            run.font.bold = (c == 1)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)

    # Set column widths
    for row in table.rows:
        row.cells[0].width = Inches(3)
        row.cells[1].width = Inches(3.5)

    doc.add_page_break()

    # ── TABLE OF CONTENTS ──
    h = doc.add_heading("Table of Contents", level=1)
    style_heading(h, 1)
    add_horizontal_line(doc)

    toc_items = [
        ("Slide 1", "Introduction — EKATVA AI", "Vipul Krishna Pandey"),
        ("Slide 2", "The Problem", "Vipul Krishna Pandey"),
        ("Slide 3", "Our Solution", "Vipul Krishna Pandey"),
        ("Slide 4", "How EKATVA Works", "Vipul Krishna Pandey"),
        ("Slide 5", "What Can EKATVA Do?", "Vipul Krishna Pandey"),
        ("Slide 6", "Technology Stack", "Vipul Krishna Pandey"),
        ("Slide 7", "Market Research", "Vipul Krishna Pandey"),
        ("Slide 8", "Why EKATVA? — Our Differentiation", "Vipul Krishna Pandey"),
        ("", "— HANDOVER TRANSITION —", ""),
        ("Slide 9", "Target Customers", "Mohan Sharma"),
        ("Slide 10", "MVP — Minimum Viable Product", "Mohan Sharma"),
        ("Slide 11", "Business Model", "Mohan Sharma"),
        ("Slide 12", "Marketing and Customer Acquisition", "Mohan Sharma"),
        ("Slide 13", "Challenges and Solutions", "Mohan Sharma"),
        ("Slide 14", "Future Roadmap", "Mohan Sharma"),
        ("Slide 15", "Conclusion", "Mohan Sharma"),
    ]

    for slide_num, title, presenter in toc_items:
        if slide_num == "":
            add_styled_paragraph(doc, f"    {title}", size=11, color=RGBColor(33, 100, 33),
                                  bold=True, italic=True, space_before=4, space_after=4)
        else:
            add_mixed_paragraph(doc, [
                {"text": f"{slide_num}:  ", "bold": True, "color": DARK_NAVY, "size": 11},
                {"text": title, "color": BLACK, "size": 11},
                {"text": f"  —  {presenter}", "color": ACCENT_BLUE, "size": 10, "italic": True},
            ], space_before=2, space_after=2, line_spacing_pt=16)

    doc.add_page_break()

    # ══════════════════════════════════════════════
    # SLIDE SCRIPTS
    # ══════════════════════════════════════════════

    slides = [
        {
            "num": 1,
            "title": "Introduction — EKATVA AI",
            "presenter": "Vipul Krishna Pandey",
            "script": "Good morning everyone. Let me start with a simple question — What if you never had to open 10 different apps to get your daily work done?\n\nThink about it. Every day, we open Gmail to check emails, Google Calendar to see our schedule, Notion to track our tasks, WhatsApp for messages, and so many more apps. We spend so much time just switching between these apps that we forget to do the actual work.\n\nThat is the problem we want to solve with EKATVA AI.\n\nEKATVA — the word itself means oneness, unity, bringing everything together. And that is exactly what our proposed product does. It is an AI-powered Personal Digital Agent that brings all your apps together under one intelligent interface.\n\nOur tagline says it perfectly — Many Apps. One Intelligence.\n\nI am Vipul Krishna Pandey, and with me is Mohan Sharma. Together, we will walk you through our startup idea today.\n\nLet us first understand the problem we are trying to solve."
        },
        {
            "num": 2,
            "title": "The Problem",
            "presenter": "Vipul Krishna Pandey",
            "script": "Before I talk about our solution, let me describe a situation that I am sure everyone here can relate to.\n\nYou wake up in the morning. You check WhatsApp — 50 messages. Then you open Gmail — 20 emails. Then you remember you have a deadline, so you open Notion. Then you check Google Calendar to see if you have any classes. Then someone sends you a document on Drive. Before you know it, you have opened 8-9 different apps and you have not actually started working yet.\n\nThis is not just a personal problem. Research shows that the average person uses 9 or more apps every single day, and we switch between apps more than a thousand times. Think about how much time that wastes.\n\nThe real issue is that these apps do not talk to each other properly. Your email does not know about your calendar. Your task manager does not know about your emails. Everything is disconnected, or fragmented.\n\nAnd the worst part? People end up spending more time managing their apps than doing actual productive work.\n\nSo what if there was one intelligent assistant that could handle all of this for you? That is exactly what EKATVA proposes to do."
        },
        {
            "num": 3,
            "title": "Our Solution",
            "presenter": "Vipul Krishna Pandey",
            "script": "Now here is our proposed solution — EKATVA AI.\n\nEKATVA is an AI-powered Personal Digital Agent. Let me break that down in simple language.\n\nInstead of opening Gmail, then Calendar, then Notion separately, you just open EKATVA and give it one simple command in plain English. For example, you could say — 'Check my emails, find anything urgent, add tasks to Notion, and organise my schedule for tomorrow.'\n\nEKATVA's AI agent would then understand what you want, break it down into smaller steps, connect to the right apps using authorised integrations, perform the tasks, and give you one final result.\n\nNow, there is an important difference I want to highlight. Most people think of AI as a chatbot — you ask a question, it gives an answer. But EKATVA is proposed as an AI agent, not just a chatbot. An AI agent can actually take actions on your behalf. It can read your emails, create calendar events, add tasks to Notion — all with your permission, of course.\n\nThe key flow is simple: One interface, one natural language command, multiple connected applications, coordinated execution, and one unified result.\n\nLet me show you exactly how this would work step by step."
        },
        {
            "num": 4,
            "title": "How EKATVA Works",
            "presenter": "Vipul Krishna Pandey",
            "script": "Let me walk you through how EKATVA would work, step by step.\n\nStep 1 — You open EKATVA, either on your phone or on the web.\nStep 2 — You type or speak a command. For example: 'Check my important emails, identify pending tasks, add them to Notion and organise my calendar for tomorrow.'\nStep 3 — The AI understands what you actually want. It is not just looking for keywords — it understands the full meaning of your instruction.\nStep 4 — It breaks this one big request into smaller tasks: read emails, find important ones, create Notion tasks, check calendar availability, create events.\nStep 5 — It identifies which connected apps it needs: Gmail, Notion, and Google Calendar.\nStep 6 — It checks if it has the right permissions to access each app.\nStep 7 — This is very important. Before doing anything sensitive, like sending an email or creating calendar events, it asks for your approval first. You always stay in control.\nStep 8 — Once approved, it executes each task one by one.\nStep 9 — Finally, it shows you a complete summary of everything it did.\n\nSo instead of spending 20-30 minutes doing all this manually across 3-4 apps, EKATVA could potentially do it in seconds with one command.\n\nNow let me show you the range of things EKATVA could do."
        },
        {
            "num": 5,
            "title": "What Can EKATVA Do?",
            "presenter": "Vipul Krishna Pandey",
            "script": "So what can EKATVA actually do? Let me divide its proposed capabilities into four areas.\n\nFirst, Personal Life. EKATVA could help you organise your daily schedule, set reminders, and manage your plans. Imagine telling it — 'Plan my day for tomorrow' — and it checks your calendar, your pending tasks, and creates an organised plan.\n\nSecond, Communication. It could summarise your important emails so you do not have to read through everything. It could even prepare draft replies and flag urgent messages.\n\nThird, Student Life — and this is very relevant for all of us here. EKATVA could track your assignments, organise your study schedule, create task lists, and make sure you never miss a deadline. Imagine saying — 'What assignments do I have this week?' — and getting a clear answer instantly.\n\nFourth, Professional Work. For working professionals, it could organise meetings, coordinate tasks across teams, prepare daily summaries, and automate repetitive work.\n\nNow, I want to be clear — these are all proposed capabilities. What EKATVA can actually do would depend on which apps are integrated and what permissions users grant. We are not claiming it can do everything from day one.\n\nNow let me talk about the technology that would make all this possible."
        },
        {
            "num": 6,
            "title": "Technology Stack",
            "presenter": "Vipul Krishna Pandey",
            "script": "Now let me explain the technology that would power EKATVA. I will keep this simple and non-technical.\n\nFor the frontend — meaning the part users would see and interact with — we plan to use React Native with Expo for the mobile app and Next.js with React for the web app. React Native lets us build one app that works on both Android and iPhone.\n\nFor the backend — the engine running behind the scenes — we would use Node.js with TypeScript and Fastify. Think of this as the brain's nervous system that processes all requests.\n\nThe AI brain itself would use models from OpenAI or Claude through their APIs. These are the same models that power ChatGPT and Claude. They help the system understand what users mean when they give natural language commands.\n\nLangGraph is a framework that helps coordinate multi-step AI agent workflows. When a user gives one command that needs 5 different steps, LangGraph helps the AI agent plan and execute each step in the right order.\n\nFor data storage, we would use PostgreSQL — a reliable, industry-standard database — with Prisma for easier database management and pgvector for storing AI-related data.\n\nFor connecting to external apps like Gmail and Calendar, we would use OAuth 2.0 for secure authentication, REST APIs for communication, and MCP — which stands for Model Context Protocol — for AI-friendly app integrations.\n\nRedis and BullMQ would handle background tasks, and we would deploy using Docker containers on cloud platforms like AWS, Railway, or Vercel.\n\nNow that you understand how we plan to build it, let us look at what already exists in the market."
        },
        {
            "num": 7,
            "title": "Market Research",
            "presenter": "Vipul Krishna Pandey",
            "script": "Before building any product, we need to understand what already exists in the market. Let me walk you through the major players.\n\nZapier is probably the most well-known automation tool. It connects over 6,000 applications and lets you create automated workflows. It is very powerful, but it requires users to manually build workflows using a visual builder. It is more of a tool for tech-savvy users.\n\nn8n is similar to Zapier but open-source and more flexible, especially with AI workflows. Again, it requires technical setup.\n\nChatGPT — I am sure everyone knows this. OpenAI has been adding connected apps and actions to ChatGPT, allowing it to interact with other services. This is actually the closest to what EKATVA proposes.\n\nLindy provides AI assistants for business workflows, and Microsoft Power Automate is Microsoft's enterprise automation platform.\n\nNow, I want to be completely honest here. These products already do many things well, and some of their capabilities overlap with what EKATVA proposes. We are not claiming they cannot do what we want to do.\n\nHowever, we see a possible opportunity in creating a simpler, more personal experience. Something designed for everyday users — students and professionals — who want a simple command interface without having to build complex workflows.\n\nThis is where I will explain how EKATVA differentiates itself."
        },
        {
            "num": 8,
            "title": "Why EKATVA? — Our Differentiation",
            "presenter": "Vipul Krishna Pandey",
            "script": "So why EKATVA? What makes our proposed approach different?\n\nLet me highlight seven key aspects of our product positioning.\n\nFirst, one command centre. Instead of 10 different apps, you have one place to manage your digital life.\nSecond, natural-language interaction. You do not need to learn how to build workflows or write automation rules. You just speak or type in plain English.\nThird, it coordinates across multiple connected apps in a single request.\nFourth, it learns your personal context and preferences over time, so it gets better at helping you.\nFifth, proactive assistance. Instead of waiting for you to ask, it can remind you about upcoming deadlines or suggest actions.\nSixth, transparent execution history. You can always see exactly what the AI did, when, and why.\nAnd seventh, user approval. Before doing anything sensitive, it always asks for your permission.\n\nA traditional chatbot just answers questions. An automation tool like Zapier requires you to manually build workflows. EKATVA's proposed approach is different — you give it a natural language command, and it understands, plans, and acts, all with your permission.\n\nAs we say — EKATVA is not just about connecting apps. It is about making your digital life easier to manage.\n\nNow I would like to invite my partner Mohan Sharma to continue the presentation. Mohan will walk you through our target customers, MVP, business model, and future plans."
        },
        # HANDOVER happens here (handled separately)
        {
            "num": 9,
            "title": "Target Customers",
            "presenter": "Mohan Sharma",
            "script": "Thank you, Vipul. Now let me tell you about who would use EKATVA.\n\nWe have identified five initial customer segments.\n\nFirst — College Students, like all of us here. Students juggle between email, calendar, assignment trackers, note-taking apps, and messaging platforms. EKATVA could help them manage all of this from one place.\n\nSecond — Working Professionals. Anyone who works with email, meetings, task management, and daily planning would benefit from a unified assistant.\n\nThird — Freelancers. Freelancers often manage multiple clients, projects, and tools. An AI assistant that coordinates across their apps could save them hours every week.\n\nFourth — Entrepreneurs. Startup founders use dozens of tools every day. EKATVA could help them stay on top of everything without getting lost in app management.\n\nFifth — Small Business Owners. They need to coordinate team tasks, handle communication, and manage schedules efficiently.\n\nOur primary initial target would be students and working professionals, because they already use multiple digital apps daily and would immediately see the value. We would start with this group and expand from there.\n\nSo how would we actually build the first version? Let me show you our MVP plan."
        },
        {
            "num": 10,
            "title": "MVP — Minimum Viable Product",
            "presenter": "Mohan Sharma",
            "script": "Now let me talk about the MVP — the Minimum Viable Product. This is the simplest, most realistic version of EKATVA that we would build first.\n\nAn MVP is not the final product. It is the first working version that has just enough features to be useful and to test whether people actually want it.\n\nOur proposed MVP would have 13 core features. Let me highlight the most important ones.\n\nUsers would be able to register and connect three apps: Gmail, Google Calendar, and Notion. These are the three apps we would start with because they cover email, scheduling, and task management — the most common daily activities.\n\nThe natural-language command interface would let users type or speak instructions in plain English. The AI would understand the intent, plan the steps, and execute them.\n\nFor example, imagine a user says: 'Check my emails, find important tasks, add them to Notion and organise my calendar for tomorrow.'\n\nHere is what the system would do:\n1. Read the user's authorised emails\n2. Use AI to identify which ones contain important tasks\n3. Prepare a task list and show it to the user\n4. Ask for approval before creating anything\n5. Once approved, create the tasks in Notion\n6. Check the user's calendar for available slots\n7. Suggest a schedule for tomorrow\n8. Display a complete summary of everything it did\n\nI want to emphasise — these are proposed features. We have not completed the full development yet, but we have a clear technical plan for how to build each feature.\n\nNow, let us talk about how EKATVA would make money."
        },
        {
            "num": 11,
            "title": "Business Model",
            "presenter": "Mohan Sharma",
            "script": "Every startup needs a business model — a clear answer to 'How will you make money?'\n\nEKATVA proposes a freemium subscription model. This means the basic version would be free, and users would pay for additional features.\n\nThe Free Plan would give users limited AI tasks per month with 2-3 connected apps. This lets people try EKATVA and see its value before paying anything.\n\nThe Pro Plan at ₹299 per month would offer more connected apps, higher task limits, and features like personal preferences and recurring workflows.\n\nThe Premium Plan at ₹699 per month would include advanced multi-step workflows, proactive assistance, and priority execution.\n\nFor businesses and teams, we would offer custom pricing with team workflows, shared automation, and administrative controls.\n\nNow, I want to be transparent — these prices are proposed assumptions for our business plan. The actual pricing would need to be validated through customer research and would depend heavily on our actual AI infrastructure costs. Running AI models costs money, and we need to make sure our pricing covers those costs while remaining affordable.\n\nThe key revenue model is: attract users with a free plan, demonstrate value, and convert them to paid plans as they rely on EKATVA more.\n\nHaving a good product is not enough. We also need a plan to reach customers."
        },
        {
            "num": 12,
            "title": "Marketing and Customer Acquisition",
            "presenter": "Mohan Sharma",
            "script": "Now let me explain how we would attract customers. We have a three-phase plan.\n\nPhase 1 is Campus Launch. We would start right here — in colleges. We would promote EKATVA through campus events, live demonstrations, and student ambassadors. Students who love the product would tell their friends about it. We would also use Instagram and LinkedIn to create awareness and run a referral programme — invite a friend, get extra free tasks.\n\nWhy start with colleges? Because students are early adopters. They try new technology quickly, and if we solve their problems, they become loyal users who carry our product into their professional lives.\n\nPhase 2 is Expand Reach. Once we have a solid student base, we would target freelancers and working professionals through productivity communities, YouTube demo videos, and user testimonials from real users.\n\nPhase 3 is Scale Up. At this stage, we would reach small businesses, offer team subscriptions, and build partnerships with relevant service providers.\n\nThe customer journey is simple: People become aware of EKATVA, they sign up for free, they experience their first successful automation, they start using it regularly, and eventually they upgrade to a paid plan because they find it valuable.\n\nOf course, every startup faces challenges. Let me be honest about ours."
        },
        {
            "num": 13,
            "title": "Challenges and Solutions",
            "presenter": "Mohan Sharma",
            "script": "Let me be honest about the challenges EKATVA would face. Every startup has challenges, and pretending they do not exist is not a good strategy.\n\nChallenge 1 — Privacy and Security. When you are accessing someone's email and calendar, security is critical. Our solution: We would use OAuth 2.0, which means users grant specific permissions without sharing passwords. All access tokens would be encrypted. We would request only the minimum permissions needed and maintain complete audit logs.\n\nChallenge 2 — AI Mistakes. AI is not perfect. What if it misunderstands a command? Our solution: Every action would show an execution preview first. Users would approve before anything sensitive happens. If something goes wrong, the execution history lets users see exactly what happened.\n\nChallenge 3 — App Restrictions. Not every app allows third-party access. Our solution: We would start only with apps that have official, supported APIs — like Gmail, Google Calendar, and Notion. We would not try to access apps that do not support integrations.\n\nChallenge 4 — AI Operating Costs. Running AI models is expensive. Our solution: Usage limits on the free plan, using efficient models where possible, and subscription revenue to cover infrastructure costs.\n\nChallenge 5 — User Trust. People are understandably cautious about giving an AI access to their accounts. Our solution: Complete transparency. Users can see everything the AI does, disconnect apps at any time, and control all permissions.\n\nOne important point — connecting an email address or phone number alone does not give EKATVA access to all apps. Each service requires its own separate, authorised integration.\n\nNow let me share our vision for the future."
        },
        {
            "num": 14,
            "title": "Future Roadmap",
            "presenter": "Mohan Sharma",
            "script": "Let me share our proposed roadmap — how we see EKATVA evolving over time.\n\nPhase 1 is the MVP, which I already described. We would build the basic AI agent with Gmail, Google Calendar, and Notion integrations. This would take about 3-4 months.\n\nPhase 2 would expand supported applications and workflows. We would add more apps like Google Drive, Slack, Trello, and others based on user demand. This would happen around months 5-8.\n\nPhase 3 is the Personal AI Manager phase. Here, we would improve how EKATVA understands individual user preferences and context. It would start offering proactive reminders and daily briefings without being asked. This would be around months 9-12.\n\nPhase 4 would explore advanced automation — potentially including browser and device automation where technically and legally supported. This is more ambitious and would be in our second year.\n\nPhase 5 is our long-term vision — building EKATVA into a comprehensive personal digital agent. A unified AI ecosystem that truly manages your entire digital life. This is a 2-3 year vision.\n\nEach phase builds on the previous one. We are not trying to build everything at once. We start small, validate, and grow.\n\nLet me conclude with our vision for the future."
        },
        {
            "num": 15,
            "title": "Conclusion",
            "presenter": "Mohan Sharma",
            "script": "And that brings us to the end of our presentation.\n\nLet me leave you with this thought:\n\nToday, we manage dozens of applications. We spend hours switching between them, organizing them, trying to make them work together. Tomorrow, we believe we should only need one intelligent assistant.\n\nEKATVA AI. Many Apps. One Intelligence.\n\nOur goal with EKATVA is not to create yet another app that people have to manage. Our goal is to create an intelligent assistant that helps them manage everything else.\n\nWe know this is an ambitious vision. We know there are strong competitors in the market. We know there are technical and business challenges to overcome. But we believe the future of personal digital management is unified, intelligent, and simple.\n\nThank you for your time and attention. We would be happy to answer any questions."
        },
    ]

    for slide in slides:
        # Add handover before slide 9
        if slide["num"] == 9:
            add_handover_box(doc)
            doc.add_page_break()

        # Slide heading
        h = doc.add_heading(f"Slide {slide['num']}: {slide['title']}", level=1)
        style_heading(h, 1, DARK_NAVY, size=18)
        add_horizontal_line(doc)

        # Presenter box
        add_presenter_box(doc, slide["presenter"])

        # Script text - split by paragraphs
        paragraphs = slide["script"].split("\n\n")
        for para_text in paragraphs:
            add_speech_text(doc, para_text)

        # Page break after each slide (except last)
        if slide["num"] < 15:
            doc.add_page_break()

    # ── END PAGE ──
    doc.add_page_break()
    for _ in range(8):
        doc.add_paragraph()
    add_styled_paragraph(doc, "— End of Speaking Script —", size=16, color=DARK_NAVY, bold=True,
                          alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    add_styled_paragraph(doc, "EKATVA AI | Many Apps. One Intelligence.", size=13, color=ACCENT_BLUE,
                          italic=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_styled_paragraph(doc, "Presenters: Vipul Krishna Pandey & Mohan Sharma", size=12, color=BLACK,
                          alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=12)

    output_path = r"D:\EKATVA\EKATVA_Presentation_Script.docx"
    doc.save(output_path)
    print(f"  ✅ Saved: {output_path}")
    return output_path


# ════════════════════════════════════════════════════════════════════
# DOCUMENT 2: FACULTY Q&A PREPARATION GUIDE
# ════════════════════════════════════════════════════════════════════

def create_qa_document():
    print("Creating Document 2: Faculty Q&A Preparation Guide...")
    doc = Document()
    set_document_margins(doc)

    # ── Style defaults ──
    style = doc.styles['Normal']
    font = style.font
    font.name = FONT_NAME
    font.size = Pt(12)
    font.color.rgb = BLACK
    style.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)

    for level in range(1, 4):
        hs = doc.styles[f'Heading {level}']
        hs.font.name = FONT_NAME
        hs.font.color.rgb = DARK_NAVY
        hs.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)

    # ── TITLE PAGE ──
    for _ in range(6):
        doc.add_paragraph()

    add_styled_paragraph(doc, "EKATVA AI", size=36, color=DARK_NAVY, bold=True,
                          alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_styled_paragraph(doc, "Faculty Q&A Preparation Guide", size=22, color=ACCENT_BLUE,
                          alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)

    add_horizontal_line(doc)

    add_styled_paragraph(doc, "Anticipated Questions and Suggested Answers", size=14, color=DARK_NAVY,
                          alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=16, space_after=8)

    add_styled_paragraph(doc, "Entrepreneurship Presentation | BTech", size=13, color=ACCENT_BLUE,
                          italic=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=8, space_after=20)

    add_horizontal_line(doc)

    add_styled_paragraph(doc, "15 Questions with Detailed Answers", size=14, color=DARK_NAVY, bold=True,
                          alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=20)

    doc.add_page_break()

    # ── Q&A PAIRS ──
    qa_pairs = [
        {
            "q": "Why did you choose this idea?",
            "a": "We chose this idea because we personally experience the problem every day. As students, we use Gmail, Google Calendar, Notion, WhatsApp, Google Drive, and many other apps daily. We noticed that we spend more time switching between these apps than actually doing productive work. When AI models like ChatGPT and Claude became powerful enough to understand natural language and take actions, we realised there was an opportunity to build an intelligent assistant that could coordinate across all these apps. The timing felt right — the AI technology is mature, the problem is real, and the market is growing."
        },
        {
            "q": "How is EKATVA different from ChatGPT?",
            "a": "ChatGPT is a general-purpose AI model that can answer questions, write content, and has started adding connected app functionality. EKATVA is proposed as a specialised Personal Digital Agent designed specifically for managing everyday tasks across multiple applications. While ChatGPT can do many things, EKATVA would focus specifically on being your personal productivity assistant with deeper integrations, transparent action history, proactive suggestions, and a user-approval system for sensitive actions. That said, we acknowledge that ChatGPT with its connected apps is one of our closest competitors, and the line between general AI assistants and specialised agents is evolving."
        },
        {
            "q": "How is EKATVA different from Zapier?",
            "a": "Zapier is an excellent automation tool that connects over 6,000 applications. However, Zapier requires users to manually create workflows using a visual builder — you need to define triggers, actions, and conditions yourself. This works great for tech-savvy users but can be overwhelming for everyday users. EKATVA proposes a natural-language approach — you just tell it what you want in plain English, and it figures out the steps. That said, Zapier is more mature, has more integrations, and is a proven product. We would be positioning EKATVA as a simpler, more personal alternative for everyday individual users."
        },
        {
            "q": "What is an AI agent?",
            "a": "An AI agent is a software program powered by artificial intelligence that can understand instructions, make plans, and take actions on behalf of the user. It is different from a simple chatbot. A chatbot just answers questions — you ask something, it responds. An AI agent can actually do things. For example, if you tell an AI agent to 'schedule a meeting with my professor tomorrow at 3 PM,' it can actually check your calendar, find the right time slot, create the event, and even send an email invitation — all with your permission. Think of it as the difference between asking someone for directions versus having a driver who takes you there."
        },
        {
            "q": "How will EKATVA access other applications?",
            "a": "EKATVA would use official APIs provided by each application. API stands for Application Programming Interface — it is a way for software programs to communicate with each other. For example, Google provides Gmail API and Calendar API that allow authorised applications to read emails and create events. We would use OAuth 2.0 for authentication, which means users log into each service themselves and grant specific permissions to EKATVA. We never see or store user passwords. Each application requires its own separate authorisation — connecting your email does not automatically give us access to your calendar or any other app. We would also explore MCP (Model Context Protocol), which is a newer standard for AI-friendly app integrations."
        },
        {
            "q": "How will user data remain safe?",
            "a": "Data security is our top priority. Here is our proposed approach: First, we would use OAuth 2.0, which means users grant limited, specific permissions without sharing passwords. Second, all access tokens would be encrypted using industry-standard encryption. Third, we would request only the minimum permissions needed — if we only need to read emails, we would not ask for permission to send emails. Fourth, we would maintain complete audit logs so users can see exactly what EKATVA accessed and when. Fifth, users can disconnect any app at any time and revoke all permissions. Sixth, we would comply with India's Digital Personal Data Protection Act and other relevant regulations. We would also conduct regular security audits as the platform grows."
        },
        {
            "q": "How will the startup earn money?",
            "a": "We propose a freemium subscription model. The basic version would be free with limited AI tasks and 2-3 connected apps. This lets users experience the value without paying. Paid plans would start at approximately ₹299 per month for the Pro plan with more apps and features, and ₹699 per month for the Premium plan with advanced capabilities. For businesses and teams, we would offer custom pricing. The key revenue strategy is: attract users with the free plan, demonstrate value through successful automation, and convert regular users to paid plans. These prices are proposed assumptions — actual pricing would be validated through market research and would depend on our real AI infrastructure costs."
        },
        {
            "q": "What is your MVP?",
            "a": "MVP stands for Minimum Viable Product — the simplest working version of our product that is useful enough to test with real users. Our proposed MVP would include: user registration, connecting Gmail, Google Calendar, and Notion, a natural-language command interface, AI understanding of instructions, multi-step task planning, email summarisation, calendar event creation, Notion task creation, execution history, and user approval before sensitive actions. We would start with just three apps because they cover the most common daily tasks: email, scheduling, and task management. The MVP would help us validate whether people actually find this useful before investing in building the full platform."
        },
        {
            "q": "What will be the initial development cost?",
            "a": "This is an estimate based on our research, not actual expenditure. For a small team of 2-3 developers working for 3-4 months:\n\n• Cloud hosting (AWS/Railway): approximately ₹5,000–15,000 per month\n• AI API costs (OpenAI/Claude): approximately ₹10,000–30,000 per month depending on usage\n• Domain and basic infrastructure: approximately ₹5,000–10,000 one-time\n• Development tools and services: approximately ₹5,000–10,000 per month\n• Total estimated MVP development cost: approximately ₹1–3 lakhs over 3-4 months\n\nThese are rough estimates. The actual costs would depend on many factors including team size, development speed, and infrastructure choices. As founding developers, our own time would be the biggest investment, which is not included in these cost estimates."
        },
        {
            "q": "What if the AI makes a mistake?",
            "a": "AI making mistakes is a real concern, and we take it very seriously. Our proposed safeguards include: First, execution previews — before performing any action, the system would show the user exactly what it plans to do. Second, user approval — for sensitive actions like sending emails or modifying calendar events, the system would always ask for explicit approval. Third, the system would never perform irreversible actions without confirmation. Fourth, a complete execution history lets users review everything the AI did. Fifth, users can undo actions where technically possible. Sixth, over time, we would use user feedback to improve the AI's accuracy. The philosophy is: the AI proposes, the user approves."
        },
        {
            "q": "Why would people pay for EKATVA?",
            "a": "People would pay for EKATVA for the same reason they pay for any productivity tool — it saves them time. If EKATVA saves a working professional even 30 minutes a day, that is 15 hours a month. At ₹299 or ₹699 per month, the value of time saved far exceeds the cost. But we recognise that people only pay for products they genuinely find useful. That is why the free plan is so important — it lets users experience the value first. If the free version does not convince them, no amount of marketing will. Our success depends entirely on building a product that people actually want to use every day."
        },
        {
            "q": "What are the major competitors?",
            "a": "Our major competitors include: Zapier (6000+ app integrations, workflow automation), n8n (open-source automation with AI), ChatGPT (AI with connected apps and actions — this is our closest competitor), Lindy (AI assistants for workflows), Microsoft Power Automate (enterprise automation), and Google Assistant and Apple Siri (basic personal assistants). We also need to watch for new AI agent startups that are emerging in this space. We respect these competitors and acknowledge that they already offer many overlapping capabilities. Our differentiation is in proposing a simpler, more personal, everyday assistant experience with transparent AI actions and strong user controls."
        },
        {
            "q": "How will you acquire customers?",
            "a": "Our customer acquisition strategy has three phases. Phase 1 — Campus Launch: Start in colleges with live demonstrations, student ambassadors, social media promotion on Instagram and LinkedIn, and a referral programme. This is low-cost and targets our primary audience. Phase 2 — Expand Reach: Target freelancers and professionals through productivity communities, YouTube demo videos, and user testimonials. Phase 3 — Scale Up: Reach small businesses with team subscriptions and partnerships. The key metric at each stage is not just signups but active usage — we need people who actually use EKATVA regularly, not just people who create an account and forget about it."
        },
        {
            "q": "What are the technical limitations?",
            "a": "We want to be transparent about our limitations. First, EKATVA can only connect to apps that provide official APIs — not all apps do. Second, AI understanding is not perfect — complex or ambiguous commands might be misinterpreted. Third, some actions have rate limits imposed by the apps themselves — for example, Gmail limits how many emails you can send per day. Fourth, real-time performance depends on AI model response times, which can vary. Fifth, some app integrations may require paid API access. Sixth, the system requires internet connectivity. Seventh, user data processing must comply with privacy regulations, which may limit certain features. We believe in being upfront about these limitations rather than making unrealistic promises."
        },
        {
            "q": "What is your long-term vision?",
            "a": "Our long-term vision for EKATVA is to become the primary way people interact with their digital world. Instead of manually managing dozens of apps, users would have one intelligent assistant that understands their preferences, anticipates their needs, and takes care of routine tasks automatically. We see this evolving in phases: first, basic task automation across a few apps; then, deeper integrations with more services; then, a truly personal AI manager that learns from your patterns; and finally, a comprehensive personal digital agent that handles everything from scheduling to communication to task management. We are not trying to replace existing apps — we are trying to make them work together intelligently. The future of digital life is unified, and we want EKATVA to be at the centre of that unification."
        },
    ]

    for i, qa in enumerate(qa_pairs):
        q_num = i + 1

        # Question number and text
        q_heading = doc.add_heading(f"Q{q_num}: {qa['q']}", level=2)
        style_heading(q_heading, 2, DARK_NAVY, size=14)

        # Answer label
        add_mixed_paragraph(doc, [
            {"text": "Answer:", "bold": True, "color": ACCENT_BLUE, "size": 11},
        ], space_before=4, space_after=4, line_spacing_pt=16)

        # Answer text - split by double newlines for multi-paragraph answers
        answer_parts = qa["a"].split("\n\n")
        for part in answer_parts:
            # Check if it's a bullet list item
            if part.startswith("•"):
                lines = part.split("\n")
                for line in lines:
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0.4)
                    p.paragraph_format.space_before = Pt(1)
                    p.paragraph_format.space_after = Pt(1)
                    p.paragraph_format.line_spacing = Pt(18)
                    run = p.add_run(line)
                    run.font.name = FONT_NAME
                    run.font.size = Pt(12)
                    run.font.color.rgb = BLACK
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
            else:
                add_styled_paragraph(doc, part, size=12, color=BLACK, space_before=2, space_after=6)

        # Separator line between Q&A pairs (except last)
        if q_num < 15:
            add_horizontal_line(doc)
            doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # ── END PAGE ──
    doc.add_page_break()
    for _ in range(8):
        doc.add_paragraph()
    add_styled_paragraph(doc, "— End of Q&A Preparation Guide —", size=16, color=DARK_NAVY, bold=True,
                          alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    add_styled_paragraph(doc, "EKATVA AI | Many Apps. One Intelligence.", size=13, color=ACCENT_BLUE,
                          italic=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_styled_paragraph(doc, "Prepared for: Entrepreneurship Presentation | BTech", size=12, color=BLACK,
                          alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=12)

    output_path = r"D:\EKATVA\EKATVA_Faculty_Questions.docx"
    doc.save(output_path)
    print(f"  ✅ Saved: {output_path}")
    return output_path


# ════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("=" * 60)
    print("  EKATVA AI — Document Generator")
    print("=" * 60)
    print()

    path1 = create_speaking_script()
    print()
    path2 = create_qa_document()

    print()
    print("=" * 60)
    print("  ✅ Both documents created successfully!")
    print(f"  📄 Script:    {path1}")
    print(f"  📄 Q&A Guide: {path2}")
    print("=" * 60)
