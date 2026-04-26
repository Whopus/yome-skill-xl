// swift-tools-version:5.9
//
// @yome/xl — iOS backend (read-only subset; bundled into the Yome iOS app).
// Like the macOS variant, real implementation lives in the iOS app
// target during the v0.1 monorepo phase; this scaffold exists for spec
// 4 layout faithfulness.

import PackageDescription

let package = Package(
    name: "XlBackendIOS",
    platforms: [.iOS(.v17)],
    products: [
        .library(name: "XlBackendIOS", targets: ["XlBackendIOS"])
    ],
    targets: [
        .target(name: "XlBackendIOS", path: "Sources/XlBackendIOS")
    ]
)
