-- xl save [--path=<save_as>]
-- If --path is empty the renderer leaves "" in saveAsPath; we branch
-- in AppleScript to pick "save as" vs plain "save".
set saveAsPath to {{path|json}}
tell application "Microsoft Excel"
    if saveAsPath is "" then
        save active workbook
    else
        save active workbook in saveAsPath
    end if
    return name of active workbook
end tell
