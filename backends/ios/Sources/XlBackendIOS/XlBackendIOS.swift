// XlBackendIOS.swift — iOS bundled backend for @yome/xl (read-only).
//
// iOS does not ship Microsoft Excel, so the iOS backend exposes only
// inspection actions (open, get, range, used, find, sheets, books).
// Real implementation is wired in the Yome iOS app target.

import Foundation

public enum XlBackendIOS {
    public static let domain = "xl"
    public static let signatureRange = ">=1.0.0 <2.0.0"
    public static let readOnlyActions: [String] = [
        "open", "books", "sheets", "sheet", "get", "range", "used", "find", "close"
    ]
}
