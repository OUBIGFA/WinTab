import pathlib, base64, zlib

B = chr(92)  # backslash
T = chr(96)  # backtick
Q = chr(34)  # double quote
md = []
a = md.append

a("# WinTab Explorer 破坏分析与修复方案")
a("")
a("> 日期：2026-04-03  ")
a("> 系统：Windows 11 25H2 (Build 26300)  ")
a("> 用户：bigfa")
a("")
a("---")
a("")
a("## 一、WinTab 如何破坏 Explorer")
a("")
a("### 1.1 核心破坏机制")
a("")
a("WinTab 通过 **注册表 Shell Verb 劫持** 和 **COM DelegateExecute 重定向** 来接管 Windows Explorer 的文件夹打开行为。这是一个深入到 Windows Shell 底层的修改，一旦出错后果严重。")
a("")
a("#### 原理图")
a("")
a(T*3)
a("")
a("### 1.2 WinTab 修改的完整注册表清单")
a("")
a("#### HKCU Shell Verb 覆盖 (9 组 × 2 值 = 18 个值)")
a("")
a("对 " + T + "Folder" + T + "、" + T + "Directory" + T + "、" + T + "Drive" + T + " 三个类，每个类的 " + T + "open" + T + "、" + T + "explore" + T + "、" + T + "opennewwindow" + T + " 三个动词：")
a("")
a("| 注册表路径 | 值名 | 被改为 |")
a("|-----------|------|--------|")
a("| " + T + "HKCU" + B + "Software" + B + "Classes" + B + "{类}" + B + "shell" + B + "{动词}" + B + "command" + T + " | " + T + "(Default)" + T + " | " + T + Q+Q + T + " (空字符串) |")
a("| " + T + "HKCU" + B + "Software" + B + "Classes" + B + "{类}" + B + "shell" + B + "{动词}" + B + "command" + T + " | " + T + "DelegateExecute" + T + " | " + T + "{FD5BF2CD-0B24-4A80-9AF3-E40F9AFC0001}" + T + " |")
