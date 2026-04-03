
import pathlib, base64, sys

# Read the markdown content from the temp file
md_path = pathlib.Path(r"E:\_Free code\WinTab	ools\_md_full.tmp")
md_content = md_path.read_text(encoding="utf-8")

# Remove BOM if present
if md_content.startswith(chr(0xFEFF)):
    md_content = md_content[1:]

# Base64 encode the content
b64 = base64.b64encode(md_content.encode("utf-8")).decode("ascii")

# Build the gen_doc.py script with the embedded base64 content
script_lines = [
    "#!/usr/bin/env python3",
    "# -*- coding: utf-8 -*-",
    '"""',
    "Generate the WinTab Explorer damage analysis and repair document.",
    "",
    r"Output: D:\Data\Desktop\WinTab-Explorer-破坏分析与修复方案.md",
    '"""',
    "",
    "import base64",
    "import pathlib",
    "",
    "",
    f"_CONTENT_B64 = """
{b64}
"""",
]

print(len(script_lines))
print("bootstrapper logic OK")
