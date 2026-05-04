-- xl headings <on|off>
set spec to {{state|json}}
set flag to true
if spec is "off" or spec is "false" or spec is "hide" then set flag to false
tell application "Microsoft Excel"
    tell active window
        set display headings to flag
    end tell
end tell
if flag then
    return "on"
else
    return "off"
end if
