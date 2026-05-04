-- xl height <row> --size=<height>
set rowNum to ({{row|json}}) as integer
set newHeight to ({{size|json}}) as real
tell application "Microsoft Excel"
    tell active sheet
        set row height of row rowNum to newHeight
    end tell
end tell
return "updated"
