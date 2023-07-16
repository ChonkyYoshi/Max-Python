def CopyParFormatting(target, source):
    target.style = source.style
    target.paragraph_format.alignment = source.paragraph_format.alignment
    target.paragraph_format.first_line_indent = \
        source.paragraph_format.first_line_indent
    target.paragraph_format.keep_together = \
        source.paragraph_format.keep_together
    target.paragraph_format.keep_with_next = \
        source.paragraph_format.keep_with_next
    target.paragraph_format.left_indent = source.paragraph_format.left_indent
    target.paragraph_format.line_spacing = source.paragraph_format.line_spacing
    target.paragraph_format.line_spacing_rule = \
        source.paragraph_format.line_spacing_rule
    target.paragraph_format.page_break_before = \
        source.paragraph_format.page_break_before
    target.paragraph_format.right_indent = source.paragraph_format.right_indent
    target.paragraph_format.space_after = source.paragraph_format.space_after
    target.paragraph_format.space_before = source.paragraph_format.space_before
    target.paragraph_format.widow_control = \
        source.paragraph_format.widow_control


def CopyRunFormatting(target, source):
    target.style = source.style
    target.font.all_caps = source.font.all_caps
    target.font.bold = source.font.bold
    target.font.color.rgb = source.font.color.rgb
    target.font.complex_script = source.font.complex_script
    target.font.cs_bold = source.font.cs_bold
    target.font.cs_italic = source.font.cs_italic
    target.font.double_strike = source.font.double_strike
    target.font.emboss = source.font.emboss
    target.font.hidden = source.font.hidden
    target.font.highlight_color = source.font.highlight_color
    target.font.imprint = source.font.imprint
    target.font.italic = source.font.italic
    target.font.math = source.font.math
    target.font.name = source.font.name
    target.font.no_proof = source.font.no_proof
    target.font.outline = source.font.outline
    target.font.rtl = source.font.rtl
    target.font.shadow = source.font.shadow
    target.font.size = source.font.size
    target.font.small_caps = source.font.small_caps
    target.font.snap_to_grid = source.font.snap_to_grid
    target.font.spec_vanish = source.font.spec_vanish
    target.font.strike = source.font.strike
    target.font.subscript = source.font.subscript
    target.font.superscript = source.font.superscript
    target.font.underline = source.font.underline
    target.font.web_hidden = source.font.web_hidden
