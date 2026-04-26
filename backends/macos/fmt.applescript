-- xl fmt <range> [--bold] [--italic] [--size] [--color] [--bg]
--                [--align] [--numfmt] [--border]
--
-- Each attribute is wrapped in (a) the dispatcher's {{#if name}}…{{/if}}
-- block — so unspecified attributes drop out at template-render time and
-- never reach AppleScript at all — and (b) a runtime try/on-error so a
-- bad value for one attribute does not abort the others.
--
-- We collect ok/fail names and return them joined as "okList|failList"
-- so the CLI can render a per-attribute summary like the Swift bridge.
set rawRef to {{range|json}}
set sheetName to ""
set addr to rawRef
if rawRef contains "@" then
    set AppleScript's text item delimiters to "@"
    set parts to text items of rawRef
    set sheetName to item 1 of parts
    set addr to item 2 of parts
    set AppleScript's text item delimiters to ""
end if

tell application "Microsoft Excel"
    if sheetName is "" then
        set targetSheet to active sheet
    else
        set targetSheet to worksheet sheetName of active workbook
    end if
    tell targetSheet
        set targetRange to range addr
        set okList to {}
        set failList to {}

        {{#if bold}}
        try
            set bold of font object of targetRange to true
            set end of okList to "bold"
        on error errMsg
            set end of failList to "bold: " & errMsg
        end try
        {{/if}}

        {{#if italic}}
        try
            set italic of font object of targetRange to true
            set end of okList to "italic"
        on error errMsg
            set end of failList to "italic: " & errMsg
        end try
        {{/if}}

        {{#if size}}
        try
            set font size of font object of targetRange to ({{size|json}} as real)
            set end of okList to "size"
        on error errMsg
            set end of failList to "size: " & errMsg
        end try
        {{/if}}

        {{#if color}}
        try
            set color of font object of targetRange to {{color|rgb}}
            set end of okList to "color"
        on error errMsg
            set end of failList to "color: " & errMsg
        end try
        {{/if}}

        {{#if bg}}
        try
            set color of interior object of targetRange to {{bg|rgb}}
            set end of okList to "bg"
        on error errMsg
            set end of failList to "bg: " & errMsg
        end try
        {{/if}}

        {{#if align}}
        try
            set alignKind to {{align|json}}
            if alignKind is "left" then
                set horizontal alignment of targetRange to horizontal align left
            else if alignKind is "right" then
                set horizontal alignment of targetRange to horizontal align right
            else if alignKind is "center" then
                set horizontal alignment of targetRange to horizontal align center
            end if
            set end of okList to "align"
        on error errMsg
            set end of failList to "align: " & errMsg
        end try
        {{/if}}

        {{#if numfmt}}
        try
            set number format of targetRange to {{numfmt|json}}
            set end of okList to "numfmt"
        on error errMsg
            set end of failList to "numfmt: " & errMsg
        end try
        {{/if}}

        {{#if border}}
        try
            set borderKind to {{border|json}}
            if borderKind is "all" then
                set borders of targetRange to {true, true, true, true}
            else if borderKind is "outline" then
                set border (border top) of targetRange to {line style:continuous, weight:weight thin}
                set border (border bottom) of targetRange to {line style:continuous, weight:weight thin}
                set border (border left) of targetRange to {line style:continuous, weight:weight thin}
                set border (border right) of targetRange to {line style:continuous, weight:weight thin}
            else if borderKind is "top" then
                set border (border top) of targetRange to {line style:continuous, weight:weight thin}
            else if borderKind is "bottom" then
                set border (border bottom) of targetRange to {line style:continuous, weight:weight thin}
            else if borderKind is "left" then
                set border (border left) of targetRange to {line style:continuous, weight:weight thin}
            else if borderKind is "right" then
                set border (border right) of targetRange to {line style:continuous, weight:weight thin}
            end if
            set end of okList to "border"
        on error errMsg
            set end of failList to "border: " & errMsg
        end try
        {{/if}}

        set AppleScript's text item delimiters to ","
        set okStr to okList as string
        set failStr to failList as string
        set AppleScript's text item delimiters to ""
        return okStr & "|" & failStr
    end tell
end tell
