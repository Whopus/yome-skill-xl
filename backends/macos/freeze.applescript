-- xl freeze [--row=<n>] [--col=<n>]
-- Freeze panes above row=<n> and to the left of col=<n>. Either may be 0 (no split that axis).
set splitRow to ({{row|json}}) as integer
set splitCol to ({{col|json}}) as integer

tell application "Microsoft Excel"
    activate
    tell active window
        try
            set freeze panes to false
        end try
        set split row to splitRow
        set split column to splitCol
        set freeze panes to true
    end tell
end tell
return "frozen"
