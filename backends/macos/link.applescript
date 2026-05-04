-- xl link <cell> --url=<url> [--text=<display>] [--tip=<tooltip>]
set rawRef to {{cell|json}}
set urlStr to {{url|json}}
set txtStr to {{text|json}}
set tipStr to {{tip|json}}
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
        set anchor to range addr
        if txtStr is "" then set txtStr to urlStr
        if tipStr is "" then
            add hyperlink anchor object anchor address urlStr text to display txtStr
        else
            add hyperlink anchor object anchor address urlStr text to display txtStr screen tip tipStr
        end if
    end tell
end tell
return "linked"
