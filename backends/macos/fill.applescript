-- xl fill <range> --values=<TSV> — write a TSV block starting at the
-- top-left of <range>. TAB separates columns; \n (literal backslash-n
-- the agent often passes in) or actual linefeed separates rows.
--
-- IMPORTANT: inside `tell application "Microsoft Excel"` the bare
-- identifier `tab` resolves to Excel's *Tab* terminology (a class name),
-- NOT the ASCII TAB character. Splitting/concatenating with `tab` from
-- inside an Excel tell block silently produces the literal string "tab"
-- and breaks TSV parsing — so we capture the real TAB into TAB_CHAR
-- before entering any Excel tell, and use TAB_CHAR everywhere after.
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)

set rawRef to {{range|json}}
set rawValues to {{values|json}}
set sheetName to ""
set addr to rawRef
if rawRef contains "@" then
    set AppleScript's text item delimiters to "@"
    set parts to text items of rawRef
    set sheetName to item 1 of parts
    set addr to item 2 of parts
    set AppleScript's text item delimiters to ""
end if

-- Normalize literal "\n" to real linefeed.
set AppleScript's text item delimiters to "\\n"
set valSegments to text items of rawValues
set AppleScript's text item delimiters to LF_CHAR
set normalized to valSegments as string
set AppleScript's text item delimiters to ""

-- Split into rows.
set AppleScript's text item delimiters to LF_CHAR
set rowList to text items of normalized
set AppleScript's text item delimiters to ""

-- Pre-split every row into columns BEFORE entering the Excel tell block,
-- because once inside `tell application "Microsoft Excel"` the `tab`
-- identifier is shadowed by Excel's terminology even when we route it
-- through TAB_CHAR (TID assignment is fine, but downstream identifiers
-- behave more predictably when the slicing happens out here).
set AppleScript's text item delimiters to TAB_CHAR
set parsedRows to {}
repeat with i from 1 to count of rowList
    set end of parsedRows to (text items of (item i of rowList))
end repeat
set AppleScript's text item delimiters to ""

tell application "Microsoft Excel"
    if sheetName is "" then
        set targetSheet to active sheet
    else
        set targetSheet to worksheet sheetName of active workbook
    end if
    tell targetSheet
        set baseRange to range addr
        set baseRow to first row index of baseRange
        set baseCol to first column index of baseRange
        repeat with i from 1 to count of parsedRows
            set colList to item i of parsedRows
            repeat with j from 1 to count of colList
                set cellVal to item j of colList
                set targetCell to cell ((baseRow + i - 1) as integer) of column ((baseCol + j - 1) as integer)
                try
                    set value of targetCell to (cellVal as real)
                on error
                    set value of targetCell to cellVal
                end try
            end repeat
        end repeat
    end tell
end tell
return "updated"
