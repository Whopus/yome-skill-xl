-- xl sheets — list sheets of active workbook as TSV: name\trows\tcols\tactive
-- See fill.applescript: bare `tab` inside an Excel tell block is Excel's
-- terminology, not the TAB character.
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)

-- Collect raw values inside the Excel tell using INDEXED access.
-- See books.applescript for why `repeat with ws in worksheets …` is unsafe
-- (contained-item references break `name of ws` reflection -> -50).
set rawNames to {}
set rawRows to {}
set rawCols to {}
set rawActive to {}
tell application "Microsoft Excel"
    set activeSheetName to (name of active sheet) as string
    set wsCount to count of worksheets of active workbook
    repeat with i from 1 to wsCount
        set wsName to (name of worksheet i of active workbook) as string
        set ur to used range of worksheet i of active workbook
        set end of rawNames to wsName
        set end of rawRows to (count of rows of ur) as string
        set end of rawCols to (count of columns of ur) as string
        if wsName is activeSheetName then
            set end of rawActive to "*"
        else
            set end of rawActive to ""
        end if
    end repeat
end tell
set sheetList to {"name" & TAB_CHAR & "rows" & TAB_CHAR & "cols" & TAB_CHAR & "active"}
repeat with i from 1 to count of rawNames
    set end of sheetList to (item i of rawNames) & TAB_CHAR & (item i of rawRows) & TAB_CHAR & (item i of rawCols) & TAB_CHAR & (item i of rawActive)
end repeat
-- Join OUTSIDE the Excel tell block: inside it, `text item delimiters`
-- collides with Excel's terminology and raises (-50) 参数错误.
set AppleScript's text item delimiters to LF_CHAR
set joined to sheetList as string
set AppleScript's text item delimiters to ""
return joined
