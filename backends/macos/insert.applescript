-- xl insert <range> [--shift=down|right]
set rawRef to {{range|json}}
set shiftSpec to {{shift|json}}
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
        if shiftSpec is "right" then
            insert into range range addr shift shift right
        else
            insert into range range addr shift shift down
        end if
    end tell
end tell
return "inserted"
