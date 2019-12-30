//
//  File.swift
//  
//
//  Created by Hugues Ferland on 2019-11-14.
//

import Foundation
import XCTest

extension XCTestCase {
    var isAttachmentAvailable: Bool {
        return ProcessInfo.processInfo.environment["XCTestConfigurationFilePath"] != nil
    }
}
