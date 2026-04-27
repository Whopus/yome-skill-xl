-- xl used — returns "sheet\trange\trows\tcols"
-- (CLI normalizes this into JSON if it wants; we keep it TSV here for parity
--  with the Swift bridge's pre-JSON output.)
-- See fill.applescript: bare `tab` inside an Excel tell block is Excel's
-- terminology, not the TAB character.
set TAB_CHAR to (ASCII character 9)

tell application "Microsoft Excel"
    set sheetName to name of active sheet
    tell active sheet
        set ur to used range
        set rowCount to count of rows of ur
        set colCount to count of columns of ur
        set addr to get address of ur
        return (sheetName as string) & TAB_CHAR & (addr as string) & TAB_CHAR & (rowCount as string) & TAB_CHAR & (colCount as string)
    end tell
end tell
