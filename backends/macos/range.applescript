-- xl range <range> — returns TSV (rows separated by linefeed, cols by tab).
-- Supports Sheet@A1:C10 syntax.
--
-- See note in fill.applescript: inside `tell application "Microsoft Excel"`
-- the bare `tab` identifier resolves to Excel's *Tab* terminology, not the
-- ASCII TAB character — concatenating with it produces the literal "tab"
-- string and corrupts every TSV consumer. Capture the real chars first.
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)

set rawRef to {{range|json}}
set sheetName to ""
set addr to rawRef
if rawRef contains "@" then
    set AppleScript's text item delimiters to "@"
    set parts to text items of rawRef
    set sheetName to item 1 of parts
    set addr to item 2 of parts
    set AppleScript's text item delimiters to ""
end if
tell application "Microsoft Excel"
    if sheetName is "" then
        set targetSheet to active sheet
    else
        set targetSheet to worksheet sheetName of active workbook
    end if
    tell targetSheet
        set r to range addr
        set rowCount to count of rows of r
        set colCount to count of columns of r
        set output to ""
        repeat with i from 1 to rowCount
            set rowData to ""
            repeat with j from 1 to colCount
                set cellVal to (value of cell i of column j of r) as string
                if j > 1 then set rowData to rowData & TAB_CHAR
                set rowData to rowData & cellVal
            end repeat
            if i > 1 then set output to output & LF_CHAR
            set output to output & rowData
        end repeat
        return output
    end tell
end tell
