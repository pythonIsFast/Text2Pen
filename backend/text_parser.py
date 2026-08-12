"""Split raw input text into text blocks and tab-separated table blocks."""


def parse_text_with_tables(text):
    lines = text.split("\n")
    blocks = []
    current_table = []
    current_text = []

    def flush_text():
        nonlocal current_text
        if current_text:
            blocks.append({"type": "text", "content": "\n".join(current_text)})
            current_text = []

    def flush_table():
        nonlocal current_table
        if current_table:
            blocks.append({"type": "table", "rows": current_table})
            current_table = []

    for line in lines:
        if "\t" in line:
            flush_text()
            current_table.append(line.split("\t"))
        else:
            flush_table()
            current_text.append(line)

    flush_text()
    flush_table()
    return blocks


def count_drawable_chars(blocks):
    total = 0
    for block in blocks:
        if block["type"] == "text":
            total += len([c for c in block["content"] if c not in (" ", "\n")])
        else:
            total += len([c for row in block["rows"] for c in "".join(row) if c != " "])
    return total


def find_unknown_chars(blocks, letter_db):
    """Return the characters used in `blocks` that have no trained strokes."""
    unknown = []
    for block in blocks:
        if block["type"] == "text":
            haystack = block["content"]
        else:
            haystack = "".join("".join(row) for row in block["rows"])
        for ch in haystack:
            if ch and ch not in (" ", "\n", "\t") and ch not in letter_db and ch not in unknown:
                unknown.append(ch)
    return unknown
