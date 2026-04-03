# -*- coding: utf-8 -*-
content = open(r'E:\_Free code\WinTab\tools\doc_content.md', 'r', encoding='utf-8').read()
with open(r'D:\Data\Desktop\WinTab-Explorer-破坏分析与修复方案.md', 'w', encoding='utf-8') as f:
    f.write(content)
print('Done')
