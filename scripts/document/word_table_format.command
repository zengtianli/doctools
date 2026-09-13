#!/bin/zsh
# Mac 版 Microsoft Word：打开目标文档，再双击本文件。
# 可选：终端传入当前 Word 文档的完整路径，用于校验处理对象。
# 维护原版：doctools/scripts/document/word_table_format.command
set -eu
if [[ "${1:-}" == "--help" || "${1:-}" == "-h" ]]; then
    print -r -- "用法：word_table_format.command [当前 Word 文档的完整路径]"
    print -r -- "Mac Word .docx：底纹无颜色、内容水平/垂直居中、内外边框 0.5 磅黑色实线；图名、表名使用文档已有的 ZDWP图名 样式。"
    print -r -- "先保存当前编辑并创建同目录 .bak-时间戳 备份，再通过 Word 文档接口整理和保存。"
    exit 0
fi
if (( $# > 1 )); then
    print -u2 -r -- "最多传入一个文档路径；查看用法：--help"
    exit 2
fi
export WORD_TABLE_LIBRARY_DIR="${0:A:h:h:h}/lib"
export WORD_TABLE_PYTHON="$(command -v python3)"
/usr/bin/osascript - "$@" <<'APPLESCRIPT'
-- 通过 Word 文档接口处理，不模拟键盘、不使用剪贴板。
on clearShading(s)
    tell application "Microsoft Word"
        set background pattern color index of s to auto
        set foreground pattern color index of s to auto
        set texture of s to texture none
    end tell
end clearShading

on solidOutline(b)
    tell application "Microsoft Word"
        set outside line style of b to line style single
        set outside line width of b to line width50 point
        set outside color index of b to black
    end tell
end solidOutline

on formatTable(t)
    tell application "Microsoft Word"
        set r to text object of t
        -- 清除单元格、段落、文字三个层面的底纹。
        my clearShading(shading of t)
        my clearShading(shading of paragraph format of r)
        my clearShading(shading of font object of r)
        set alignment of paragraph format of r to align paragraph center
        set vertical alignment of every cell of r to cell align vertical center
        -- 外框及内部横竖线统一为 0.5 磅黑色单实线。
        set b to border options of t
        my solidOutline(b)
        if (has horizontal of b) or (has vertical of b) then
            set inside line style of b to line style single
            set inside line width of b to line width50 point
            set inside color index of b to black
        end if
        -- 单元格直接设置的无边框/虚线会覆盖整表边框，必须逐格统一。
        -- 按真实 cell 遍历，避免合并单元格的行列寻址问题。
        repeat with c in (get cells of r)
            my solidOutline(border options of c)
        end repeat
        set tableTotal to 1
        repeat with innerTable in (get tables of t)
            set tableTotal to tableTotal + my formatTable(innerTable)
        end repeat
        return tableTotal
    end tell
end formatTable

on captionPlan(sourcePath)
    set planner to ¬
        "import sys, io, zipfile" & linefeed & ¬
        "from xml.etree import ElementTree as E" & linefeed & ¬
        "sys.path.insert(0, sys.argv[1])" & linefeed & ¬
        "from caption_re import parse, BID_STRICT_CAPTION, is_fig_caption_style, is_table_caption_style" & linefeed & ¬
        "w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'" & linefeed & ¬
        "with zipfile.ZipFile(sys.argv[2]) as z:" & linefeed & ¬
        "    root = E.fromstring(z.read('word/document.xml'))" & linefeed & ¬
        "    styles = E.fromstring(z.read('word/styles.xml'))" & linefeed & ¬
        "names = {s.get(w+'styleId'): s.find(w+'name').get(w+'val') for s in styles.findall(w+'style') if s.find(w+'name') is not None}" & linefeed & ¬
        "body = root.find(w+'body')" & linefeed & ¬
        "parents = {c: p for p in body.iter() for c in p}" & linefeed & ¬
        "ordered = []" & linefeed & ¬
        "def visit(el):" & linefeed & ¬
        "    for child in el:" & linefeed & ¬
        "        visit(child)" & linefeed & ¬
        "    if el.tag in (w+'p', w+'tr'):" & linefeed & ¬
        "        ordered.append(el)" & linefeed & ¬
        "visit(body)" & linefeed & ¬
        "captions = []" & linefeed & ¬
        "for index, p in enumerate(ordered, 1):" & linefeed & ¬
        "    if p.tag != w+'p':" & linefeed & ¬
        "        continue" & linefeed & ¬
        "    text = ''.join(t.text or '' for t in p.iter(w+'t')).strip()" & linefeed & ¬
        "    if not parse(text, BID_STRICT_CAPTION):" & linefeed & ¬
        "        continue" & linefeed & ¬
        "    ancestor = parents.get(p)" & linefeed & ¬
        "    in_cell = False" & linefeed & ¬
        "    while ancestor is not None:" & linefeed & ¬
        "        in_cell |= ancestor.tag in (w+'tc', w+'txbxContent')" & linefeed & ¬
        "        ancestor = parents.get(ancestor)" & linefeed & ¬
        "    if in_cell:" & linefeed & ¬
        "        continue" & linefeed & ¬
        "    style = p.find(w+'pPr/'+w+'pStyle')" & linefeed & ¬
        "    name = names.get(style.get(w+'val'), '') if style is not None else ''" & linefeed & ¬
        "    siblings = list(parents[p])" & linefeed & ¬
        "    pos = siblings.index(p)" & linefeed & ¬
        "    before = next((e for e in reversed(siblings[:pos]) if e.tag in (w+'p', w+'tbl') and (e.tag == w+'tbl' or ''.join(e.itertext()).strip() or any(e.iter(w+'drawing')) or any(e.iter(w+'pict')))), None)" & linefeed & ¬
        "    after = next((e for e in siblings[pos+1:] if e.tag in (w+'p', w+'tbl') and (e.tag == w+'tbl' or ''.join(e.itertext()).strip())), None)" & linefeed & ¬
        "    near_table = text.startswith('表') and after is not None and after.tag == w+'tbl'" & linefeed & ¬
        "    near_image = text.startswith('图') and before is not None and (any(before.iter(w+'drawing')) or any(before.iter(w+'pict')))" & linefeed & ¬
        "    if name == 'ZDWP图名' or is_fig_caption_style(name) or is_table_caption_style(name) or near_table or near_image:" & linefeed & ¬
        "        if any(c in text for c in ('\\r', '\\n', '\\t')):" & linefeed & ¬
        "            raise SystemExit('题注含制表符或段落换行，需人工确认：' + text)" & linefeed & ¬
        "        captions.append((index, text))" & linefeed & ¬
        "# Word 在每个表格行末额外计一个段落；调用方同时核对总数和每条题注全文。" & linefeed & ¬
        "print(len(ordered))" & linefeed & ¬
        "for index, text in captions:" & linefeed & ¬
        "    print(str(index).zfill(10) + '|' + text)"
    return do shell script (quoted form of (system attribute "WORD_TABLE_PYTHON")) & " -c " & quoted form of planner & " " & quoted form of (system attribute "WORD_TABLE_LIBRARY_DIR") & " " & quoted form of sourcePath
end captionPlan

on captionTargets(d, planText)
    set planLines to paragraphs of planText
    tell application "Microsoft Word"
        if (count of paragraphs of d) is not (item 1 of planLines as integer) then error "Word 段落结构与已保存文档不一致，未执行本次格式修改。"
        set targets to {}
        repeat with lineNumber from 2 to count of planLines
            set entryText to item lineNumber of planLines
            set paragraphIndex to (text 1 thru 10 of entryText) as integer
            set expectedText to text 12 thru -1 of entryText
            set p to paragraph paragraphIndex of d
            set actualText to content of text object of p
            -- 全文精确核对，避免使用 Word 与 XML 不一致的段落编号写错位置。
            if actualText is not (expectedText & return) then error "题注定位不一致，未执行本次格式修改：" & expectedText
            set end of targets to {paragraphIndex, expectedText}
        end repeat
        return targets
    end tell
end captionTargets

on applyCaptions(d, targets)
    tell application "Microsoft Word"
        repeat with target in targets
            set r to text object of paragraph (item 1 of target) of d
            if content of r is not ((item 2 of target) & return) then error "文档在运行期间发生变化，已停止题注修改。"
            set style of r to "ZDWP图名"
            reset paragraph format of r
            reset font object of r
        end repeat
        return count of targets
    end tell
end applyCaptions

on run argv
    if application "Microsoft Word" is not running then error "请先在 Microsoft Word 中打开要整理的文档。"
    with timeout of 600 seconds
        tell application "Microsoft Word"
            if (count of documents) is 0 then error "请先打开要整理的 Word 文档。"
            set d to active document
            set sourcePath to posix full name of d
            if sourcePath does not start with "/" then error "请先把文档保存到本机，再运行本脚本。"
            if sourcePath does not end with ".docx" then error "本脚本用于 .docx 文档，请先保存为 .docx。"
            if (count of argv) > 0 then
                if sourcePath is not item 1 of argv then error "当前 Word 文档与指定路径不一致，尚未修改。"
            end if
            if read only of d then error "当前文档是只读状态，无法修改。"
            if protection type of d is not no document protection then error "当前文档有编辑保护，请解除保护后再运行。"
            try
                get Word style "ZDWP图名" of d
            on error
                error "文档中没有已有的 ZDWP图名 样式，请先准备该样式；本次未修改。"
            end try
            -- 先保存当前编辑内容，再备份；绝不在磁盘上覆盖 Word 的内存文档。
            save d
            set stamp to do shell script "/bin/date +%Y%m%d-%H%M%S"
            set backupPath to sourcePath & ".bak-" & stamp
            do shell script "if /bin/test -e " & quoted form of backupPath & "; then exit 1; fi; /bin/cp -p " & quoted form of sourcePath & " " & quoted form of backupPath
            set targets to my captionTargets(d, my captionPlan(sourcePath))
            set totalTables to 0
            -- 正文、页眉页脚、文本框、脚注等区域；嵌套表格由 formatTable 递归处理。
            set storyKinds to {main text story, footnotes story, endnotes story, comments story, text frame story, even pages header story, primary header story, even pages footer story, primary footer story, first page header story, first page footer story}
            repeat with storyKind in storyKinds
                set currentRange to get story range d story type storyKind
                repeat
                    if currentRange is missing value or currentRange is "" then exit repeat
                    -- Word 对不存在的脚注等区域有时返回无效引用，而不是 missing value。
                    try
                        set rangeTables to get tables of currentRange
                    on error errMessage number errNumber
                        if errNumber is -1728 then exit repeat
                        error errMessage number errNumber
                    end try
                    repeat with t in rangeTables
                        set totalTables to totalTables + my formatTable(t)
                    end repeat
                    try
                        set currentRange to next story range of currentRange
                    on error errMessage number errNumber
                        if errNumber is -1728 then exit repeat
                        error errMessage number errNumber
                    end try
                end repeat
            end repeat
            set totalCaptions to my applyCaptions(d, targets)
            save d
            return "已完成：" & totalTables & " 个表格，底纹无颜色，文字水平、垂直居中，内外边框为 0.5 磅黑色实线；" & totalCaptions & " 处图名、表名已套用 ZDWP图名 样式。" & linefeed & "文档：" & sourcePath & linefeed & "备份：" & backupPath
        end tell
    end timeout
end run
APPLESCRIPT
