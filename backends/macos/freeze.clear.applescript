-- xl freeze.clear — unfreeze panes on the active window
tell application "Microsoft Excel"
    tell active window
        try
            set freeze panes to false
        end try
        try
            set split row to 0
            set split column to 0
        end try
    end tell
end tell
return "cleared"
