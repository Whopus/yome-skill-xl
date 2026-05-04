-- xl filter <range> [--field=<n>] [--criteria1=<v>] [--criteria2=<v>] [--op=and|or]
-- Toggles auto filter on the given range; with --field/--criteria1 also applies a filter.
set rawRef to {{range|json}}
set fieldNum to {{field|json}}
set crit1 to {{criteria1|json}}
set crit2 to {{criteria2|json}}
set opSpec to {{op|json}}
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
        if fieldNum is "" then
            -- Just toggle auto filter on the range (header row).
            try
                auto filter targetRange
            on error
                auto filter range addr
            end try
        else
            set fNum to (fieldNum as integer)
            set opVal to filter values
            if opSpec is "and" then set opVal to filter and
            if opSpec is "or" then set opVal to filter or
            if crit2 is not "" then
                auto filter targetRange field fNum criteria1 crit1 operator opVal criteria2 crit2
            else if crit1 is not "" then
                auto filter targetRange field fNum criteria1 crit1
            else
                auto filter targetRange field fNum
            end if
        end if
    end tell
end tell
return "filtered"
