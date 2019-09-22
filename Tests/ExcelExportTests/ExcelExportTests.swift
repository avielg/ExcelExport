/**
 *  ExcelExport
 *
 *  Copyright (c) 2016 Aviel Gross. Licensed under the MIT license, as follows:
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software"), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in all
 *  copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 *  SOFTWARE.
 */

import Foundation
import XCTest
@testable import ExcelExport

class ExcelExportTests: XCTestCase {
    
    func testCellValue() {
        let cell = ExcelCell("15")
        
        XCTAssertEqual(cell.value, "15")
    }
    
    func testCellAttribute() {
        let cell = ExcelCell("", [TextAttribute.font([TextAttribute.FontStyle.bold,TextAttribute.FontStyle.color(Color.blue)])])
        XCTAssertEqual(TextAttribute.styleValue(for: cell.attributes), "<Font ss:Bold=\"1\" ss:Color=\"#0000FF\"/>")
    }
    
    func testRow() {
        let cell = ExcelCell("33")
        let row = ExcelRow([cell])
        
        XCTAssertEqual(row.cells.count, 1)
    }
    
    func testSheet() {
        let cell = ExcelCell("53")
        let row = ExcelRow([cell])
        let sheet = ExcelSheet([row], name: "Test")
        
        XCTAssertEqual(sheet.rows.count, 1)
        XCTAssertEqual(sheet.name, "Test")
    }
    
    func testExport() {
        let expectation = self.expectation(description: "Async creation of Excel file.")
        var exportResultCalled: Bool = false

        // arrange
        let cells = [ExcelCell("Age : "), ExcelCell("50", [TextAttribute.backgroundColor(Color.yellow), TextAttribute.font([TextAttribute.FontStyle.bold])])]
        let sheet = ExcelSheet([ExcelRow(cells), ExcelRow(cells), ExcelRow(cells)], name: "Test")
        let sheet2 = ExcelSheet([ExcelRow(cells)], name: "Test2")
        
        // act
        ExcelExport.export([sheet,sheet2], fileName: "test") { url in
            print("done function called with \(String(describing: url))")
            if let file = url {
                do {
                    exportResultCalled = try file.checkResourceIsReachable()
                    print("... the file \(file) exist and is reachable \(exportResultCalled).")
                }
                catch {
                    print("... url not reachable (not a file on the local system.")
                }
            }
            else {
                print("... no file produced.")
            }
            expectation.fulfill()
        }
        
        // assert
        waitForExpectations(timeout: 1)
        XCTAssertTrue(exportResultCalled, "Le fichier n'a pas été créé.")
    }
}

#if os(Linux)
extension ExcelExportTests {
    static var allTests : [(String, (ExcelExportTests) -> () throws -> Void)] {
        return [
            ("testExample", testExample),
        ]
    }
}
#endif
