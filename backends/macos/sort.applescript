-- xl sort <range> --by=<col1>[,<col2>[,<col3>]] [--order=asc|desc[,asc|desc[,asc|desc]]] [--header=true|false]
-- col is a column letter (A, B, AA...) within the range OR an absolute column letter.
set rawRef to {{range|json}}
set byCols to {{by|json}}
set orderSpec to {{order|json}}
set hasHeader to {{header|json}}
set sheetName to ""
set addr to rawRef
if rawRef contains "@" then
    set AppleScript's text item delimiters to "@"
    set parts to text items of rawRef
    set sheetName to item 1 of parts
    set addr to item 2 of parts
    set AppleScript's text item delimiters to ""
end if

-- Pre-parse columns/orders OUTSIDE the Excel tell block.
set AppleScript's text item delimiters to ","
set colList to text items of byCols
set orderList to text items of orderSpec
set AppleScript's text item delimiters to ""

set headerArg to header guess
if hasHeader is "true" then
    set headerArg to header yes
else if hasHeader is "false" then
    set headerArg to header no
end if

tell application "Microsoft Excel"
    if sheetName is "" then
        set targetSheet to active sheet
    else
        set targetSheet to worksheet sheetName of active workbook
    end if
    tell targetSheet
        set targetRange to range addr
        set k1 to missing value
        set k2 to missing value
        set k3 to missing value
        set o1 to sort ascending
        set o2 to sort ascending
        set o3 to sort ascending
        if (count of colList) >= 1 then
            set k1 to range ((item 1 of colList) & "1")
        end if
        if (count of colList) >= 2 then
            set k2 to range ((item 2 of colList) & "1")
        end if
        if (count of colList) >= 3 then
            set k3 to range ((item 3 of colList) & "1")
        end if
        if (count of orderList) >= 1 and (item 1 of orderList) is "desc" then set o1 to sort descending
        if (count of orderList) >= 2 and (item 2 of orderList) is "desc" then set o2 to sort descending
        if (count of orderList) >= 3 and (item 3 of orderList) is "desc" then set o3 to sort descending
        if k1 is missing value then
            error "xl sort: missing --by"
        end if
        if k3 is not missing value then
            sort targetRange key1 k1 order1 o1 key2 k2 order2 o2 key3 k3 order3 o3 header headerArg
        else if k2 is not missing value then
            sort targetRange key1 k1 order1 o1 key2 k2 order2 o2 header headerArg
        else
            sort targetRange key1 k1 order1 o1 header headerArg
        end if
    end tell
end tell
return "sorted"
