-- xl get <cell> — returns "value\tformula"
-- Supports Sheet@A1 syntax via on-the-fly @-split.
--
-- See fill.applescript: bare `tab` inside `tell application "Microsoft Excel"`
-- is Excel's terminology, not the TAB character. Always use TAB_CHAR.
set TAB_CHAR to (ASCII character 9)

set rawRef to {{cell|json}}
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
        set cellVal to value of cell addr
        set cellFormula to formula of cell addr
        return (cellVal as string) & TAB_CHAR & cellFormula
    end tell
end tell
