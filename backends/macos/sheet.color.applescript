-- xl sheet.color [--name=<sheet>] --color=<color>  (color: empty|none clears)
set sName to {{name|json}}
set colorSpec to {{color|json}}
set rgbVal to {{color|rgb}}

tell application "Microsoft Excel"
    if sName is "" then
        set targetWS to active sheet
    else
        set targetWS to worksheet sName of active workbook
    end if
    if colorSpec is "" or colorSpec is "none" then
        try
            set tab color index of targetWS to no color index
        end try
    else
        try
            set color of tab of targetWS to rgbVal
        on error
            try
                set tab color of targetWS to rgbVal
            end try
        end try
    end if
end tell
return "tab color"
