-- xl set <cell> --value=<v> [--type=text|number|auto] | --formula=<f>
set rawRef to {{cell|json}}
set theValue to {{value|json}}
set theFormula to {{formula|json}}
set theType to {{type|json}}
set sheetName to ""
set addr to rawRef
if rawRef contains "@" then
    set AppleScript's text item delimiters to "@"
    set parts to text items of rawRef
    set sheetName to item 1 of parts
    set addr to item 2 of parts
    set AppleScript's text item delimiters to ""
end if
if theValue is "" and theFormula is "" then
    error "xl set: missing --value or --formula"
end if
tell application "Microsoft Excel"
    if sheetName is "" then
        set targetSheet to active sheet
    else
        set targetSheet to worksheet sheetName of active workbook
    end if
    tell targetSheet
        if theFormula is not "" then
            set formula of cell addr to theFormula
        else
            -- value branch
            if theType is "text" then
                set number format of cell addr to "@"
                set value of cell addr to theValue
            else if theType is "number" then
                try
                    set value of cell addr to (theValue as real)
                on error
                    error "--type=number but value '" & theValue & "' is not numeric"
                end try
            else
                -- auto: try numeric, fall back to string
                try
                    set value of cell addr to (theValue as real)
                on error
                    set value of cell addr to theValue
                end try
            end if
        end if
    end tell
end tell
return "updated"
