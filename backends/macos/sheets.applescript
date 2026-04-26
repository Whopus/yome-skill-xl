-- xl sheets — list sheets of active workbook as TSV: name\trows\tcols\tactive
tell application "Microsoft Excel"
    set sheetList to {"name" & tab & "rows" & tab & "cols" & tab & "active"}
    set activeSheetName to name of active sheet
    repeat with ws in worksheets of active workbook
        set wsName to name of ws
        set isActive to ""
        if wsName is activeSheetName then set isActive to "*"
        set ur to used range of ws
        set rowCount to count of rows of ur
        set colCount to count of columns of ur
        set end of sheetList to wsName & tab & (rowCount as string) & tab & (colCount as string) & tab & isActive
    end repeat
    set AppleScript's text item delimiters to linefeed
    return sheetList as string
end tell
