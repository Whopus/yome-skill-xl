-- xl view <kind>  kind: normal|page-break|page-layout
set kindSpec to {{kind|json}}
tell application "Microsoft Excel"
    tell active window
        if kindSpec is "page-break" then
            set view to page break preview
        else if kindSpec is "page-layout" then
            try
                set view to page layout view
            end try
        else
            set view to normal view
        end if
    end tell
end tell
return kindSpec
