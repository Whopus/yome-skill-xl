-- xl printarea <range>  (use empty string or "clear" to clear)
set rawRef to {{range|json}}
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
    if addr is "" or addr is "clear" then
        set print area of page setup object of targetSheet to ""
    else
        set print area of page setup object of targetSheet to addr
    end if
end tell
return "printarea"
