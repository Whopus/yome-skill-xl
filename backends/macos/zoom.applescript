-- xl zoom <percent>  (10..400)
set z to ({{percent|json}}) as integer
tell application "Microsoft Excel"
    tell active window
        set zoom to z
    end tell
end tell
return (z as string)
