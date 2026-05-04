-- xl group <range> [--axis=row|col]  (default row)
set rawRef to {{range|json}}
set axisKind to {{axis|json}}
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
        set targetRange to range addr
        if axisKind is "col" then
            group entire column of targetRange
        else
            group entire row of targetRange
        end if
    end tell
end tell
return "grouped"
