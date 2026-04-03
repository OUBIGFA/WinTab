import pathlib
p = pathlib.Path(r"E:\_Free code\WinTab	ools\_md_output.txt")
f = open(p, "w", encoding="utf-8", newline="
")
f.write("test content with backslash: HKCU\Software\Classes
")
f.close()
print("done")
