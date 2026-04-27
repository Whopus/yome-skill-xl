-- xl sheets — list sheets of active workbook as TSV: name\trows\tcols\tactive
-- See fill.applescript: bare `tab` inside an Excel tell block is Excel's
-- terminology, not the TAB character.
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)

tell application "Microsoft Excel"
    set sheetList to {"name" & TAB_CHAR & "rows" & TAB_CHAR & "cols" & TAB_CHAR & "active"}
    set activeSheetName to name of active sheet
    repeat with ws in worksheets of active workbook
        set wsName to name of ws
        set isActive to ""
        if wsName is activeSheetName then set isActive to "*"
        set ur to used range of ws
        set rowCount to count of rows of ur
        set colCount to count of columns of ur
        set end of sheetList to wsName & TAB_CHAR & (rowCount as string) & TAB_CHAR & (colCount as string) & TAB_CHAR & isActive
    end repeat
    set AppleScript's text item delimiters to LF_CHAR
    return sheetList as string
end tell
