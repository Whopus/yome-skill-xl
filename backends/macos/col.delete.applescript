-- xl col.delete <column> [--count=1]
set colLetter to {{column|json}}
set howMany to ({{count|json}}) as integer
tell application "Microsoft Excel"
    tell active sheet
        if howMany is 1 then
            set targetRange to range (colLetter & ":" & colLetter)
        else
            set startCol to first column index of range (colLetter & "1")
            set endColIdx to startCol + howMany - 1
            set targetRange to range ((colLetter & ":" & (address of cell 1 of column endColIdx)) as string)
        end if
        delete range targetRange shift shift left
    end tell
end tell
return "deleted " & howMany & " column(s) starting at " & colLetter
