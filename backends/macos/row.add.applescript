-- xl row.add <row> [--count=1]
set startRow to ({{row|json}}) as integer
set howMany to ({{count|json}}) as integer
tell application "Microsoft Excel"
    tell active sheet
        set endRow to startRow + howMany - 1
        set targetRange to range (startRow & ":" & endRow as string)
        insert into range targetRange shift shift down
    end tell
end tell
return "inserted " & howMany & " row(s) at " & startRow
