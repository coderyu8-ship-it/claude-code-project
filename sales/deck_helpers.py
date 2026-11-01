def font(run, size, bold=False, color=INK):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rpr = run._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set('typeface', FONT)


def text(s, x, y, w, h, lines, size=14, bold=False, color=INK,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.25):
    """lines: str | list of str | list of (str, size, bold, color)."""
    tb = s.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(lines, str):
        lines = [lines]
    for i, ln in enumerate(lines):
        if isinstance(ln, str):
            ln = (ln, size, bold, color)
        t, sz, b, c = ln
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = t
        font(r, sz, b, c)
    return tb


def box(s, x, y, w, h, fill=WHITE, line=None, shape=MSO_SHAPE.RECTANGLE, lw=1):
    b = s.shapes.add_shape(shape, Emu(x), Emu(y), Emu(w), Emu(h))
    b.shadow.inherit = False
    if fill is None:
        b.fill.background()
    else:
        b.fill.solid()
        b.fill.fore_color.rgb = fill
    if line is None:
        b.line.fill.background()
    else:
        b.line.color.rgb = line
        b.line.width = Pt(lw)
    return b


def pic(s, path, x, y, w=None, h=None, border=False):
    iw, ih = Image.open(path).size
    if w and not h:
        h = int(w * ih / iw)
    if h and not w:
        w = int(h * iw / ih)
    p = s.shapes.add_picture(path, Emu(x), Emu(y), Emu(w), Emu(h))
    if border:
        p.line.color.rgb = LINE
        p.line.width = Pt(1)
    return p


def notes(s, body):
    s.notes_slide.notes_text_frame.text = body.strip()


