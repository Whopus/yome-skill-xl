-- xl autofill --src=<range> --dest=<range> [--type=default|copy|series|formats|values]
set srcRef to {{src|json}}
set destRef to {{dest|json}}
set tSpec to {{type|json}}

set ftype to fill default
if tSpec is "copy" then set ftype to fill copy
if tSpec is "series" then set ftype to fill series
if tSpec is "formats" then set ftype to fill formats
if tSpec is "values" then set ftype to fill values

set sheetName to ""
set sAddr to srcRef
if srcRef contains "@" then
    set AppleScript's text item delimiters to "@"
    set parts to text items of srcRef
    set sheetName to item 1 of parts
    set sAddr to item 2 of parts
    set AppleScript's text item delimiters to ""
end if
set dAddr to destRef
if destRef contains "@" then
    set AppleScript's text item delimiters to "@"
    set dParts to text items of destRef
    set dAddr to item 2 of dParts
    set AppleScript's text item delimiters to ""
end if

tell application "Microsoft Excel"
    if sheetName is "" then
        set targetSheet to active sheet
    else
        set targetSheet to worksheet sheetName of active workbook
    end if
    tell targetSheet
        autofill range sAddr destination range dAddr type ftype
    end tell
end tell
return "autofilled"
