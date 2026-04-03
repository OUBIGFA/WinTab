import pathlib
import base64
import textwrap

# The markdown content, base64-encoded to avoid any escaping issues.
# We will generate this encoding and embed it.

# Step 1: Build the markdown content
md_lines = []
