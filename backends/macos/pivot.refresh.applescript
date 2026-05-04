-- xl pivot.refresh [--name=<pt_name>]  (omit name to refresh all on active sheet)
set ptName to {{name|json}}
tell application "Microsoft Excel"
    tell active sheet
        if ptName is "" then
            repeat with pt in pivot tables
                try
                    refresh pt
                end try
            end repeat
            return "refreshed all"
        else
            refresh pivot table ptName
            return "refreshed"
        end if
    end tell
end tell
