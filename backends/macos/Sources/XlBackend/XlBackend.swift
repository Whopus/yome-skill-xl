// XlBackend.swift — macOS bundled backend for @yome/xl.
//
// Real implementation currently lives in the Yome macOS app at
// Yome/macOS/Bridge/ExcelBridge.swift. This file is a placeholder so the
// skill scaffold compiles standalone if someone opens it via SwiftPM
// (spec 4 standard structure).

import Foundation

public enum XlBackend {
    public static let domain = "xl"
    public static let signatureRange = ">=1.0.0 <2.0.0"

    /// In-app integration is wired in PromptBuilder + YomeAgentEngine; this
    /// type just declares that the backend exists for the registry.
    public struct DispatchRequest {
        public let action: String
        public let positionals: [String]
        public let flags: [String: String]
        public init(action: String, positionals: [String], flags: [String: String]) {
            self.action = action
            self.positionals = positionals
            self.flags = flags
        }
    }
}
