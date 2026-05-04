-- xl validation <range> --kind=list|number|text|date
--   list:   --values=v1,v2,v3 (comma-separated)
--   number: --op=greater|less|between|equal --v1=<n> [--v2=<n>]
--   date:   --op=greater|less|between|equal --v1=<d> [--v2=<d>]
--   text:   --op=length-greater|length-less|length-between --v1=<n> [--v2=<n>]
-- [--prompt=str] [--errorTitle=str] [--errorMessage=str]
set rawRef to {{range|json}}
set kindSpec to {{kind|json}}
set valsSpec to {{values|json}}
set opSpec to {{op|json}}
set v1 to {{v1|json}}
set v2 to {{v2|json}}
set promptStr to {{prompt|json}}
set eTitle to {{errorTitle|json}}
set eMsg to {{errorMessage|json}}

set sheetName to ""
set addr to rawRef
if rawRef contains "@" then
    set AppleScript's text item delimiters to "@"
    set parts to text items of rawRef
    set sheetName to item 1 of parts
    set addr to item 2 of parts
    set AppleScript's text item delimiters to ""
end if

set vType to validate list
if kindSpec is "number" then set vType to validate whole number
if kindSpec is "decimal" then set vType to validate decimal
if kindSpec is "date" then set vType to validate date
if kindSpec is "text" then set vType to validate text length

set vOp to between
if opSpec is "greater" or opSpec is "length-greater" then set vOp to greater
if opSpec is "less" or opSpec is "length-less" then set vOp to less
if opSpec is "equal" then set vOp to equal

tell application "Microsoft Excel"
    if sheetName is "" then
        set targetSheet to active sheet
    else
        set targetSheet to worksheet sheetName of active workbook
    end if
    tell targetSheet
        set targetRange to range addr
        try
            delete validation targetRange
        end try
        if kindSpec is "list" then
            add validation targetRange type validate list operator between formula1 ("=" & valsSpec)
        else if v2 is not "" then
            add validation targetRange type vType operator vOp formula1 v1 formula2 v2
        else
            add validation targetRange type vType operator vOp formula1 v1
        end if
        try
            if promptStr is not "" then set input message of validation of targetRange to promptStr
        end try
        try
            if eTitle is not "" then set error title of validation of targetRange to eTitle
            if eMsg is not "" then set error message of validation of targetRange to eMsg
        end try
    end tell
end tell
return "validation"
