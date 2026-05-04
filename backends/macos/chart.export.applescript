-- xl chart.export <name> --path=<file.png> [--format=png|jpg|gif]
set chartName to {{name|json}}
set outPath to {{path|json}}
set fmt to {{format|json}}
if fmt is "" then set fmt to "png"

tell application "Microsoft Excel"
    tell active sheet
        try
            set theChart to chart of chart object chartName
        on error
            error "chart not found: " & chartName
        end try
        export theChart in outPath as fmt
    end tell
end tell
return outPath
