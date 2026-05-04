-- xl filter.clear — remove all filters on the active sheet
tell application "Microsoft Excel"
    tell active sheet
        try
            show all data
        end try
        try
            set auto filter mode of active sheet to false
        end try
    end tell
end tell
return "cleared"
