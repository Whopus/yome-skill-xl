# @yome/xl

Excel workbook editing commands for Yome agents (info / open / new / save / close /
books / sheets / sheet / sheet.add / sheet.rename / sheet.delete / get / range
/ used / find / set / fill / clear / fmt / width / row.* / col.* / merge /
unmerge / export). One of the official skills.

## Layout (spec v0.1, section 4)

```
yome-skill-xl/
├── yome-skill.json                       manifest (slug / domain / delivery / capabilities)
├── README.md
├── signature/
│   └── xl.signature.json                 LLM-facing command signature (truth)
├── backends/
│   ├── macos/                            Declarative manifest + .applescript templates,
│   │                                     consumed by cli/src/skills/runner/dispatcher.ts.
│   │                                     Also includes a Swift scaffold for the future
│   │                                     bundled-into-app delivery (spec 8.5).
│   ├── ios/                              Swift module bundled into Yome iOS app (read-only subset)
│   ├── node/                             TS backend stub (Phase 2)
│   └── sandbox/                          TS state machine for hub replays / benchmarks
├── viewer/
│   └── index.html                        single-direction trace renderer
├── cases/                                community-contributed Replays
└── benchmarks/                           officially scored cases
```

## Status during v0.1 monorepo phase

The signature in `signature/xl.signature.json` is byte-aligned with the
runtime descriptor `Server/agent/commands/xlCommands.ts`. The macOS
implementation currently lives inside the Yome app target as
`Yome/macOS/Bridge/ExcelBridge.swift`; `backends/macos/` ships an OTA-style
declarative manifest + AppleScript templates so the **CLI hub-skill
dispatcher** (`cli/src/skills/runner/dispatcher.ts`) can run it without
needing the bundled app — that's what makes Excel installable as a hub
skill from the CLI.

The compress functions live in `backends/sandbox/src/compress.ts`
(spec 4.4 location).

When the skill is split out of the monorepo (spec 8.5), the Swift sources
will be `git mv`d into `backends/macos/Sources/XlBackend/`.
