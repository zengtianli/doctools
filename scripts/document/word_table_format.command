#!/bin/zsh
# Mac 版 Microsoft Word：打开目标文档，再双击本文件。
# 可选：终端传入当前 Word 文档的完整路径，用于校验处理对象。
# 维护原版：doctools/scripts/document/word_table_format.command
set -eu
if [[ "${1:-}" == "--help" || "${1:-}" == "-h" ]]; then
    print -r -- "用法：word_table_format.command [当前 Word 文档的完整路径]"
    print -r -- "Mac Word .docx：底纹无颜色、内容水平/垂直居中、内外边框 0.5 磅黑色实线。"
    print -r -- "先保存当前编辑并创建同目录 .bak-时间戳 备份，再通过 Word 文档接口整理和保存。"
    exit 0
fi
if (( $# > 1 )); then
    print -u2 -r -- "最多传入一个文档路径；查看用法：--help"
    exit 2
fi
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
            -- 先保存当前编辑内容，再备份；绝不在磁盘上覆盖 Word 的内存文档。
            save d
            set stamp to do shell script "/bin/date +%Y%m%d-%H%M%S"
            set backupPath to sourcePath & ".bak-" & stamp
            do shell script "if /bin/test -e " & quoted form of backupPath & "; then exit 1; fi; /bin/cp -p " & quoted form of sourcePath & " " & quoted form of backupPath
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
            save d
            return "已完成：" & totalTables & " 个表格，底纹无颜色，文字水平、垂直居中，内外边框为 0.5 磅黑色实线。" & linefeed & "文档：" & sourcePath & linefeed & "备份：" & backupPath
        end tell
    end timeout
end run
APPLESCRIPT
