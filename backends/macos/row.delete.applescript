-- xl row.delete <row> [--count=1]
set startRow to ({{row|json}}) as integer
set howMany to ({{count|json}}) as integer
tell application "Microsoft Excel"
    tell active sheet
        set endRow to startRow + howMany - 1
        set targetRange to range (startRow & ":" & endRow as string)
        delete range targetRange shift shift up
    end tell
end tell
return "deleted " & howMany & " row(s) starting at " & startRow
