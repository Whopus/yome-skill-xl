-- xl new [<path>]
-- Creates a new blank workbook. If --path is given, also save as path.
tell application "Microsoft Excel"
    activate
    delay 1
    make new workbook
    delay 0.5
    {{#if path}}
    save active workbook in {{path|json}}
    {{/if}}
    return name of active workbook
end tell
