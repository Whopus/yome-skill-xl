-- xl wrap <range> [--state=on|off]
set rawRef to {{range|json}}
set spec to {{state|json}}
set flag to true
if spec is "off" or spec is "false" then set flag to false
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
        set wrap text of range addr to flag
    end tell
end tell
if flag then
    return "wrapped"
else
    return "unwrapped"
end if
