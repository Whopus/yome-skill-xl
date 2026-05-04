-- xl sheet.hide <name> [--state=hidden|very-hidden|visible]
set sName to {{name|json}}
set stateSpec to {{state|json}}
if stateSpec is "" then set stateSpec to "hidden"

set vState to sheet hidden
if stateSpec is "very-hidden" then set vState to sheet very hidden
if stateSpec is "visible" then set vState to sheet visible

tell application "Microsoft Excel"
    set visible of worksheet sName of active workbook to vState
end tell
return stateSpec
