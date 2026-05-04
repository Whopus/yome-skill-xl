-- xl paste <dest> [--what=all|values|formulas|formats|comments] [--transpose=true|false]
set rawRef to {{dest|json}}
set whatKind to {{what|json}}
set transposeFlag to {{transpose|json}}
set sheetName to ""
set addr to rawRef
if rawRef contains "@" then
    set AppleScript's text item delimiters to "@"
    set parts to text items of rawRef
    set sheetName to item 1 of parts
    set addr to item 2 of parts
    set AppleScript's text item delimiters to ""
end if

set pasteType to paste all
if whatKind is "values" then set pasteType to paste values
if whatKind is "formulas" then set pasteType to paste formulas
if whatKind is "formats" then set pasteType to paste formats
if whatKind is "comments" then set pasteType to paste comments

set tflag to false
if transposeFlag is "true" then set tflag to true

tell application "Microsoft Excel"
    if sheetName is "" then
        set targetSheet to active sheet
    else
        set targetSheet to worksheet sheetName of active workbook
    end if
    tell targetSheet
        try
            paste special destination range addr what pasteType with transpose
        on error
            paste special destination range addr what pasteType
        end try
    end tell
end tell
return "pasted"
