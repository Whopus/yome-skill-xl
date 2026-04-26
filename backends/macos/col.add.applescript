-- xl col.add <column> [--count=1]
-- For multi-column inserts we expand the letter range as "A:B" etc.
-- For count=1 we just use "A:A".
set colLetter to {{column|json}}
set howMany to ({{count|json}}) as integer
tell application "Microsoft Excel"
    tell active sheet
        if howMany is 1 then
            set targetRange to range (colLetter & ":" & colLetter)
        else
            -- Compute the trailing column letter via Excel's own COLUMN() helper.
            set startCol to first column index of range (colLetter & "1")
            set endColIdx to startCol + howMany - 1
            set targetRange to range ((colLetter & ":" & (address of cell 1 of column endColIdx)) as string)
        end if
        insert into range targetRange shift shift right
    end tell
end tell
return "inserted " & howMany & " column(s) at " & colLetter
