-- xl books — list open workbooks as TSV: name\tsheets
-- See fill.applescript: bare `tab` inside an Excel tell block is Excel's
-- terminology, not the TAB character.
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)

-- Collect raw values inside the Excel tell using INDEXED access.
-- `repeat with wb in workbooks` yields a contained-item reference that
-- breaks `name of wb` reflection on some Excel builds (-50 参数错误);
-- `workbook i` is the safe form.
set rawNames to {}
set rawCounts to {}
tell application "Microsoft Excel"
    set wbCount to count of workbooks
    repeat with i from 1 to wbCount
        set end of rawNames to (name of workbook i) as string
        set end of rawCounts to (count of worksheets of workbook i) as string
    end repeat
end tell
set bookList to {"name" & TAB_CHAR & "sheets"}
repeat with i from 1 to count of rawNames
    set end of bookList to (item i of rawNames) & TAB_CHAR & (item i of rawCounts)
end repeat
-- Join OUTSIDE the Excel tell block: inside it, `text item delimiters`
-- collides with Excel's terminology and raises (-50) 参数错误.
set AppleScript's text item delimiters to LF_CHAR
set joined to bookList as string
set AppleScript's text item delimiters to ""
return joined
