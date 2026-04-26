-- xl books — list open workbooks as TSV: name\tsheets
tell application "Microsoft Excel"
    set bookList to {"name" & tab & "sheets"}
    repeat with wb in workbooks
        set wbName to name of wb
        set sheetCount to count of worksheets of wb
        set end of bookList to wbName & tab & (sheetCount as string)
    end repeat
    set AppleScript's text item delimiters to linefeed
    return bookList as string
end tell
