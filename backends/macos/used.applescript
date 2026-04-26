-- xl used — returns "sheet\trange\trows\tcols"
-- (CLI normalizes this into JSON if it wants; we keep it TSV here for parity
--  with the Swift bridge's pre-JSON output.)
tell application "Microsoft Excel"
    tell active sheet
        set ur to used range
        set rowCount to count of rows of ur
        set colCount to count of columns of ur
        set addr to get address of ur
        set sheetName to name of active sheet
        return sheetName & tab & addr & tab & (rowCount as string) & tab & (colCount as string)
    end tell
end tell
