-- xl autofit <range> [--axis=col|row|both]  (default col)
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
if axisKind is "" then set axisKind to "col"

tell application "Microsoft Excel"
    if sheetName is "" then
        set targetSheet to active sheet
    else
        set targetSheet to worksheet sheetName of active workbook
    end if
    tell targetSheet
        set targetRange to range addr
        if axisKind is "col" or axisKind is "both" then
            try
                autofit (entire column of targetRange)
            end try
        end if
        if axisKind is "row" or axisKind is "both" then
            try
                autofit (entire row of targetRange)
            end try
        end if
    end tell
end tell
return "autofit"
