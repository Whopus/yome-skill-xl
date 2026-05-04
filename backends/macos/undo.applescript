-- xl undo
tell application "Microsoft Excel"
    try
        undo
    end try
end tell
return "undo"
