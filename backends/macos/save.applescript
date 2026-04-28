-- xl save [--path=<save_as>]
-- If --path is empty the renderer leaves "" in saveAsPath; we branch
-- in AppleScript to pick "save as" vs plain "save".
set saveAsPath to {{path|json}}
tell application "Microsoft Excel"
    if saveAsPath is "" then
        save active workbook
    else
        set display alerts to false
        try
            save active workbook in saveAsPath
            set display alerts to true
        on error errMsg number errNum
            set display alerts to true
            error errMsg number errNum
        end try
    end if
    return name of active workbook
end tell
