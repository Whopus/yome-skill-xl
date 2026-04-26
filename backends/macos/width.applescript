-- xl width <column> --size=<width>
-- column is a letter (A, B, AA, ...). We resolve it via Excel's own
-- range "A:A" notation rather than computing a numeric index ourselves.
set colLetter to {{column|json}}
set newWidth to {{size|json}}
tell application "Microsoft Excel"
    tell active sheet
        set column width of range (colLetter & ":" & colLetter) to (newWidth as real)
    end tell
end tell
return "updated"
