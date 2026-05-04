-- xl chart.delete <name>  (chart object name; use "*" to delete all on active sheet)
set chartName to {{name|json}}
tell application "Microsoft Excel"
    tell active sheet
        if chartName is "*" then
            set delCount to 0
            repeat with co in chart objects
                try
                    delete co
                    set delCount to delCount + 1
                end try
            end repeat
            return "deleted " & (delCount as string)
        else
            try
                delete chart object chartName
                return "deleted"
            on error
                error "chart not found: " & chartName
            end try
        end if
    end tell
end tell
