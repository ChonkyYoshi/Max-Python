from docx.text.run import Run
import itertools
from copy import deepcopy


def isolate_run(paragraph, start, end):
    """Return docx.text.Run object containing only `paragraph.text[start:end]`.

    Runs are split as required to produce a new run at the `start` that ends
    at `end`.
    Runs are unchanged if the indicated range of text already occupies its
    own run. The resulting run object is returned.

    `start` and `end` are as in Python slice notation. For example, the first
    three characters of the paragraph have (start, end) of (0, 3). `end` is
    *not* the index of the last character. These correspond to `match.start()`
    and `match.end()` of a regex match object and `s[start:end]` in Python
    slice notation.
    """
    rs = tuple(paragraph._p.r_lst)

    def advance_to_run_containing_start(start, end):
        """Return (r_idx, start, end) triple indicating start run and adjusted
        offsets.

        The start run is the run the `start` offset occurs in. The returned
        `start` and
        `end` values are adjusted to be relative to the start of `r_idx`.
        """
        # --- add 0 at end so `r_ends[-1] == 0` ---
        r_ends = tuple(itertools.accumulate(len(r.text) for r in rs)) + (0,)
        r_idx = 0
        while start >= r_ends[r_idx]:
            r_idx += 1
        skipped_rs_offset = r_ends[r_idx - 1]
        return rs[r_idx], r_idx, start - skipped_rs_offset, end -\
            skipped_rs_offset

    def split_off_prefix(r, start, end):
        """Return adjusted `end` after splitting prefix off into separate run.

        Does nothing if `r` is already the start of the isolated run.
        """
        if start > 0:
            prefix_r = deepcopy(r)
            r.addprevious(prefix_r)
            r.text = r.text[start:]
            prefix_r.text = prefix_r.text[:start]
        return end - start

    def split_off_suffix(r, end):
        """Split `r` at `end` such that suffix is in separate following run."""
        suffix_r = deepcopy(r)
        r.addnext(suffix_r)
        r.text = r.text[:end]
        suffix_r.text = suffix_r.text[end:]

    def lengthen_run(r, r_idx, end):
        """Add prefixes of following runs to `r` until `end` is reached."""
        while len(r.text) < end:
            suffix_len_reqd = end - len(r.text)
            r_idx += 1
            next_r = rs[r_idx]
            if len(next_r.text) <= suffix_len_reqd:
                # --- subsume next run ---
                r.text = r.text + next_r.text
                next_r.getparent().remove(next_r)
                continue
            if len(next_r.text) > suffix_len_reqd:
                # --- take prefix from next run ---
                r.text = r.text + next_r.text[:suffix_len_reqd]
                next_r.text = next_r.text[suffix_len_reqd:]

    # --- 1. skip over any runs before the one containing the start of our
    # range ---
    r, r_idx, start, end = advance_to_run_containing_start(start, end)

    # --- 2. split first run where our range starts, placing "prefix" to our
    # range
    # ---    in a new run inserted just before this run. After this, our run
    # will begin
    # ---    at the right point and the left-hand side of our work is done.
    end = split_off_prefix(r, start, end)

    # --- 3. if run is longer than isolation-range we need to split-off a
    # suffix run ---
    if len(r.text) > end:
        split_off_suffix(r, end)

    # --- 4. But if our run is shorter than the desired isolation-range we need
    # to
    # ---    lengthen it by taking text from subsequent runs
    elif len(r.text) < end:
        lengthen_run(r, r_idx, end)

    # --- if neither 3 nor 4 apply it's because the run already ends in the
    # right place
    # --- and there's no further work to be done.

    return Run(r, paragraph)


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
