-- xl new [<path>]
-- Creates a new blank workbook. If --path is given, also save as path.
tell application "Microsoft Excel"
    activate
    delay 1
    make new workbook
    delay 0.5
    {{#if path}}
    set display alerts to false
    try
        save active workbook in {{path|json}}
        set display alerts to true
    on error errMsg number errNum
        set display alerts to true
        error errMsg number errNum
    end try
    {{/if}}
    return name of active workbook
end tell
