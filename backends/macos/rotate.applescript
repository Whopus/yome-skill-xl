-- xl rotate <range> --degrees=<-90..90>
set rawRef to {{range|json}}
set deg to ({{degrees|json}}) as integer
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
        set orientation of range addr to deg
    end tell
end tell
return (deg as string)
