-- xl fill <range> --values=<TSV> — write a TSV block starting at the
-- top-left of <range>. TAB separates columns; \n (literal backslash-n
-- the agent often passes in) or actual linefeed separates rows.
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
set AppleScript's text item delimiters to linefeed
set normalized to valSegments as string
set AppleScript's text item delimiters to ""

-- Split into rows then columns.
set AppleScript's text item delimiters to linefeed
set rowList to text items of normalized
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
        repeat with i from 1 to count of rowList
            set rowText to item i of rowList
            set AppleScript's text item delimiters to tab
            set colList to text items of rowText
            set AppleScript's text item delimiters to ""
            repeat with j from 1 to count of colList
                set cellVal to item j of colList
                try
                    set value of cell ((baseRow + i - 1) as integer) of column ((baseCol + j - 1) as integer) to (cellVal as real)
                on error
                    set value of cell ((baseRow + i - 1) as integer) of column ((baseCol + j - 1) as integer) to cellVal
                end try
            end repeat
        end repeat
    end tell
end tell
return "updated"
