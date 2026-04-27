-- xl books — list open workbooks as TSV: name\tsheets
-- See fill.applescript: bare `tab` inside an Excel tell block is Excel's
-- terminology, not the TAB character.
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)

tell application "Microsoft Excel"
    set bookList to {"name" & TAB_CHAR & "sheets"}
    repeat with wb in workbooks
        set wbName to name of wb
        set sheetCount to count of worksheets of wb
        set end of bookList to wbName & TAB_CHAR & (sheetCount as string)
    end repeat
    set AppleScript's text item delimiters to LF_CHAR
    return bookList as string
end tell
