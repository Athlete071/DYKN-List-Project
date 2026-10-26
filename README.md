from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.text import MSO_AUTO_SIZE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Clean, consistent palette
NAVY = RGBColor(25, 43, 74)
BLUE = RGBColor(49, 94, 168)
TEAL = RGBColor(37, 139, 132)
PALE = RGBColor(239, 244, 250)
WHITE = RGBColor(255, 255, 255)
DARK = RGBColor(45, 55, 72)
GREY = RGBColor(103, 116, 134)
ORANGE = RGBColor(225, 142, 63)
GREEN = RGBColor(48, 128, 83)
RED = RGBColor(175, 72, 72)

def set_bg(slide, color=PALE):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_text(slide, x, y, w, h, text, size=20, color=DARK, bold=False, font="Aptos",
             align=None, valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    if align is not None:
        p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return box

def add_rect(slide, x, y, w, h, color, radius=False, line_color=None):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.color.rgb = line_color or color
    return sh

def add_header(slide, title, subtitle=None, num=None):
    add_rect(slide, 0, 0, 13.333, 0.18, TEAL)
    add_text(slide, 0.65, 0.48, 12.0, 0.55, title, 27, NAVY, True)
    if subtitle:
        add_text(slide, 0.68, 1.08, 12.0, 0.45, subtitle, 13, GREY)
    if num:
        add_text(slide, 12.25, 7.08, 0.5, 0.2, str(num), 9, GREY, align=PP_ALIGN.RIGHT)

def add_bullet_list(slide, x, y, w, h, items, size=19, color=DARK, gap=12):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.02)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.font.name = "Aptos"
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.space_after = Pt(gap)
        p.bullet = True
    return box

# 1 Cover
s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s, NAVY)
add_rect(s, 0, 0, 0.22, 7.5, TEAL)
add_text(s, 0.9, 1.0, 11.4, 0.45, "HACKTOBERFEST HACK DAY • PROJECT PRESENTATION", 13, RGBColor(147, 210, 205), True)
add_text(s, 0.9, 1.9, 11.2, 1.2, "DYKN-List-Project", 36, WHITE, True)
add_text(s, 0.95, 3.1, 10.6, 0.8, "A beginner-friendly list management project concept", 22, RGBColor(221, 230, 243))
add_rect(s, 0.95, 4.45, 5.6, 1.0, RGBColor(39, 62, 98), True)
add_text(s, 1.2, 4.68, 5.1, 0.5, "Project status: Initial repository created", 16, WHITE, True)
add_text(s, 0.95, 6.55, 11.0, 0.35, "GitHub: github.com/Athlete071/DYKN-List-Project", 13, RGBColor(201, 214, 233))
add_text(s, 12.25, 7.08, 0.5, 0.2, "1", 9, RGBColor(201,214,233), align=PP_ALIGN.RIGHT)

# 2 Overview
s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s); add_header(s, "1. Project overview", "A simple idea for learning how to build and manage lists.", 2)
add_rect(s, 0.7, 1.8, 5.7, 3.8, WHITE, True)
add_text(s, 1.0, 2.1, 5.0, 0.45, "The idea", 21, BLUE, True)
add_text(s, 1.0, 2.7, 4.9, 2.2, "DYKN-List-Project is planned as a small application where a user can create and manage a list of items in one place.", 19, DARK)
add_rect(s, 6.8, 1.8, 5.7, 3.8, WHITE, True)
add_text(s, 7.1, 2.1, 5.0, 0.45, "Who it may help", 21, TEAL, True)
add_bullet_list(s, 7.1, 2.7, 4.9, 2.5, ["Beginners learning basic app development", "People who want a simple list view", "A first project to practise Git and GitHub"], 17, DARK, 14)
add_text(s, 0.85, 6.15, 11.7, 0.45, "Note: This describes the proposed idea, not completed functionality.", 14, ORANGE, True)

# 3 Problem and solution
s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s); add_header(s, "2. Problem & proposed solution", "The goal is to make a basic list easy to keep organized.", 3)
add_rect(s, 0.75, 1.85, 5.65, 3.7, WHITE, True)
add_text(s, 1.05, 2.15, 4.9, 0.4, "Problem to explore", 21, BLUE, True)
add_bullet_list(s, 1.05, 2.75, 4.8, 2.3, ["Notes and items can become scattered", "It can be hard to see all items together", "A beginner needs a small, clear project scope"], 18, DARK, 16)
add_rect(s, 6.85, 1.85, 5.65, 3.7, WHITE, True)
add_text(s, 7.15, 2.15, 4.9, 0.4, "Proposed solution", 21, TEAL, True)
add_bullet_list(s, 7.15, 2.75, 4.8, 2.3, ["Show list items on one screen", "Let users add and remove items", "Keep the interface simple and beginner-friendly"], 18, DARK, 16)
add_text(s, 0.85, 6.15, 11.7, 0.45, "These are planned features; they are not implemented yet.", 14, ORANGE, True)

# 4 Implemented vs proposed
s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s); add_header(s, "3. Current status: implemented vs proposed", "This slide is intentionally transparent about the repository today.", 4)
add_rect(s, 0.75, 1.8, 5.65, 4.45, RGBColor(232, 246, 237), True)
add_text(s, 1.05, 2.1, 4.9, 0.45, "Implemented so far", 21, GREEN, True)
add_bullet_list(s, 1.05, 2.75, 4.8, 2.8, ["Public GitHub repository created", "Repository name: DYKN-List-Project", "README.md added", "Initial commit exists"], 18, DARK, 16)
add_rect(s, 6.85, 1.8, 5.65, 4.45, RGBColor(255, 245, 229), True)
add_text(s, 7.15, 2.1, 4.9, 0.45, "Proposed / not built yet", 21, ORANGE, True)
add_bullet_list(s, 7.15, 2.75, 4.8, 2.9, ["Add a new list item", "Delete or edit an item", "Save items between visits", "Improve layout and accessibility"], 18, DARK, 16)
add_text(s, 0.85, 6.55, 11.7, 0.3, "No working app or live demo is being claimed at this stage.", 13, RED, True)

# 5 Proposed features
s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s); add_header(s, "4. Planned features", "A realistic first version that can be built step by step.", 5)
features = [
    ("01", "Add an item", "Type an item and add it to the list."),
    ("02", "View items", "See all saved items in a clear list."),
    ("03", "Edit / delete", "Correct an item or remove it."),
    ("04", "Save list", "Explore browser storage or a simple database."),
]
for i, (n, title, desc) in enumerate(features):
    x = 0.8 + (i % 2) * 6.15
    y = 1.85 + (i // 2) * 2.15
    add_rect(s, x, y, 5.65, 1.65, WHITE, True)
    add_rect(s, x+0.22, y+0.25, 0.72, 0.72, TEAL, True)
    add_text(s, x+0.22, y+0.41, 0.72, 0.25, n, 14, WHITE, True, align=PP_ALIGN.CENTER)
    add_text(s, x+1.15, y+0.25, 4.1, 0.35, title, 19, NAVY, True)
    add_text(s, x+1.15, y+0.75, 4.15, 0.65, desc, 15, GREY)
add_text(s, 0.85, 6.45, 11.7, 0.35, "All four features are proposed future work, not completed features.", 13, ORANGE, True)

# 6 Technology / learning plan
s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s); add_header(s, "5. Beginner learning plan", "The technology stack can be chosen after confirming the event requirements.", 6)
steps = [
    ("Step 1", "Learn the basics", "Understand HTML, CSS and JavaScript or the required stack."),
    ("Step 2", "Build a tiny version", "Create a page that adds and displays list items."),
    ("Step 3", "Test it", "Try empty inputs, duplicates and deleting items."),
    ("Step 4", "Use GitHub", "Commit changes and explain what each change does."),
]
for i,(label,title,desc) in enumerate(steps):
    y = 1.75 + i*1.15
    add_rect(s, 0.9, y, 1.45, 0.65, BLUE if i<2 else TEAL, True)
    add_text(s, 0.97, y+0.16, 1.3, 0.25, label, 14, WHITE, True, align=PP_ALIGN.CENTER)
    add_text(s, 2.7, y+0.03, 2.7, 0.35, title, 18, NAVY, True)
    add_text(s, 5.35, y+0.03, 6.9, 0.62, desc, 16, DARK)
add_text(s, 0.9, 6.55, 11.6, 0.35, "This is a proposed learning plan; no technical stack is claimed as implemented.", 13, ORANGE, True)

# 7 Why participate
s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s); add_header(s, "6. What I want to learn at Hack Day", "A beginner's goals for participating and contributing.", 7)
add_rect(s, 0.85, 1.85, 11.65, 3.9, WHITE, True)
add_bullet_list(s, 1.25, 2.25, 10.7, 2.9, [
    "Learn how to turn a small idea into a working project",
    "Practise GitHub basics: commits, issues and pull requests",
    "Get feedback on the project scope and next steps",
    "Understand open-source collaboration and good documentation",
], 20, DARK, 18)
add_text(s, 0.9, 6.2, 11.5, 0.45, "Goal: make honest progress, ask questions and learn from the community.", 15, TEAL, True)

# 8 Next steps / submission note
s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s, NAVY)
add_rect(s, 0, 0, 13.333, 0.18, TEAL)
add_text(s, 0.8, 0.65, 11.8, 0.6, "7. Next steps", 28, WHITE, True)
add_text(s, 0.85, 1.5, 11.4, 0.45, "Before submitting the form", 20, RGBColor(147, 210, 205), True)
add_bullet_list(s, 0.9, 2.15, 11.3, 2.9, [
    "Check the event's rules: a presentation alone may not guarantee venue eligibility.",
    "Keep the repository link public and make sure the README explains the project idea.",
    "Upload this presentation only if a proposal-stage presentation is accepted.",
    "Do not claim features, teammates or registration details that are not true.",
], 18, WHITE, 16)
add_rect(s, 0.9, 5.65, 11.5, 0.8, RGBColor(39, 62, 98), True)
add_text(s, 1.15, 5.87, 11.0, 0.35, "Repository: github.com/Athlete071/DYKN-List-Project", 17, WHITE, True)
add_text(s, 12.25, 7.08, 0.5, 0.2, "8", 9, RGBColor(201,214,233), align=PP_ALIGN.RIGHT)

out = "/mnt/data/DYKN_List_Project_Presentation.pptx"
prs.save(out)
print(out)
# DYKN-List-Project
