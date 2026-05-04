-- xl shape.add --kind=rectangle|oval|line|arrow|text [--left=N] [--top=N] [--width=N] [--height=N] [--text=str]
set kindSpec to {{kind|json}}
set leftPos to ({{left|json}}) as real
set topPos to ({{top|json}}) as real
set widthVal to ({{width|json}}) as real
set heightVal to ({{height|json}}) as real
set txtStr to {{text|json}}

if widthVal = 0 then set widthVal to 100
if heightVal = 0 then set heightVal to 60

set shapeKind to rectangle
if kindSpec is "oval" then set shapeKind to oval
if kindSpec is "line" then set shapeKind to line
if kindSpec is "arrow" then set shapeKind to right arrow
if kindSpec is "text" then set shapeKind to text box

tell application "Microsoft Excel"
    tell active sheet
        set sh to make new shape at end with properties {auto shape type:shapeKind, left position:leftPos, top:topPos, width:widthVal, height:heightVal}
        if txtStr is not "" then
            try
                set content of text frame of sh to txtStr
            on error
                try
                    set text of text frame of sh to txtStr
                end try
            end try
        end if
        return name of sh
    end tell
end tell
