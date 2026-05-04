-- xl indent <range> --level=<0..15>
set rawRef to {{range|json}}
set lvl to ({{level|json}}) as integer
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
        set indent level of range addr to lvl
    end tell
end tell
return (lvl as string)
