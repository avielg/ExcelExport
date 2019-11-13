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
    
    func testMergeDown1Row() {
        // arrange
        let row1 = ExcelRow([ExcelCell("merge down 1", [], rowspan: 1), ExcelCell("first row cell 2")])
        let row2 = ExcelRow([ExcelCell("second row cell 2"), ExcelCell("second row cell 3")])
        let sheet = ExcelSheet([row1, row2], name: "Sheet1")
        
        // act
        let (exportResultCalled, url) = export([sheet])
        
        // assert
        XCTAssertTrue(exportResultCalled, "No file created.")
        XCTAssertEqual(valueOn(url, row: 2, cell: 1), "2", "Oups, cell index on second row should be 2.")
        XCTAssertEqual(valueOn(url, row: 2, cell: 2), "3", "Oups, cell index on second row should be 2.")
    }

    func testMergeDown2Rows() {
        let row1 = ExcelRow([ExcelCell("merge down 2", [], rowspan: 2), ExcelCell("first row cell 2")])
        let row2 = ExcelRow([ExcelCell("second row cell 2")])
        let row3 = ExcelRow([ExcelCell("third row cell 2")])
        let sheet = ExcelSheet([row1, row2, row3], name: "Sheet1")
        
        // act
        let (exportResultCalled, url) = export([sheet])
        
        // assert
        XCTAssertTrue(exportResultCalled, "No file created.")
        XCTAssertEqual(valueOn(url, row: 3, cell: 1), "2", "Oups, first cell index on third row should be 2.")
    }

    func testMergeDownOnThe2ndColumn() {
        // arrange
        let row1 = ExcelRow([ExcelCell("cell1 row1"),ExcelCell("cell2 row1 mergedown 1", [], rowspan: 1), ExcelCell("cell3 row1")])
        let row2 = ExcelRow([ExcelCell("cell1 row2"),ExcelCell("cell3 row2"), ExcelCell("cell4 row2")])
        let sheet = ExcelSheet([row1, row2], name: "Sheet1")
        
        // act
        let (exportResultCalled, url) = export([sheet])
        
        // assert
        XCTAssertTrue(exportResultCalled, "No file created.")
        XCTAssertEqual(valueOn(url, row: 2, cell: 2), "3", "Oups, second cell index on second row should be 3.")
        XCTAssertEqual(valueOn(url, row: 2, cell: 3), "4", "Oups, third cell index on second row should be 4.")
    }
    
    func testMergeDownResetBetweenSheet() {
        let row1 = ExcelRow([ExcelCell("cell1 row1 mergedown 2", [], rowspan: 2),ExcelCell("cell2 row1"),ExcelCell("cell3 row1 mergedown 2", [], rowspan: 2),ExcelCell("cell4 row1")])
        let row2 = ExcelRow([ExcelCell("cell1 row2"),ExcelCell("cell2 row2")])
        let sheet1 = ExcelSheet([row1,row2], name: "Sheet1")
        let row3 = ExcelRow([ExcelCell("cell1 row3"),ExcelCell("cell2 row3")])
        let sheet2 = ExcelSheet([row3], name: "Sheet2")
        
        let (exportResultCalled, url) = export([sheet1, sheet2])
        
        XCTAssertTrue(exportResultCalled, "No file created.")
        XCTAssertEqual(valueOn(url, row: 2, cell: 1), "2")
        XCTAssertEqual(valueOn(url, row: 2, cell: 2), "4")
        XCTAssertEqual(valueOn(url, sheet: 2, row: 1, cell: 1), nil)
    }
    
    func testMergeDownIn2Columns() {
        let row1 = ExcelRow([ExcelCell("cell1 row1 mergedown 2", [], rowspan: 2),ExcelCell("cell2 row1"),ExcelCell("cell3 row1 mergedown 2", [], rowspan: 2),ExcelCell("cell4 row1")])
        let row2 = ExcelRow([ExcelCell("cell1 row2"),ExcelCell("cell2 row2")])
        let row3 = ExcelRow([ExcelCell("cell1 row3"),ExcelCell("cell2 row3")])
        let sheet = ExcelSheet([row1,row2,row3], name: "Sheet1")
        
        let (exportResultCalled, url) = export([sheet])
        
        XCTAssertTrue(exportResultCalled, "No file created.")
        XCTAssertEqual(valueOn(url, row: 2, cell: 1), "2")
        XCTAssertEqual(valueOn(url, row: 2, cell: 2), "4")
        XCTAssertEqual(valueOn(url, row: 3, cell: 1), "2")
    }
    
    func testMergeDownOnTwoSubsequentRowBlock() {
        let row1 = ExcelRow([ExcelCell("cell1 row1 mergedown 2", [], rowspan: 2),ExcelCell("cell2 row1")])
        let row2 = ExcelRow([ExcelCell("cell1 row2")])
        let row3 = ExcelRow([ExcelCell("cell1 row3")])
        let row4 = ExcelRow([ExcelCell("cell1 row4 mergedown 2", [], rowspan: 2),ExcelCell("cell2 row4")])
        let row5 = ExcelRow([ExcelCell("cell1 row5")])
        let row6 = ExcelRow([ExcelCell("cell1 row6")])
        let sheet1 = ExcelSheet([row1,row2,row3,row4,row5,row6], name: "Sheet1")
        
        let (exportResultCalled, url) = export([sheet1])
        
        XCTAssertTrue(exportResultCalled, "No file created.")
        XCTAssertEqual(valueOn(url, row: 2, cell: 1), "2")
        XCTAssertEqual(valueOn(url, row: 4, cell: 1), nil)
        XCTAssertEqual(valueOn(url, row: 5, cell: 1), "2")
    }
    
    func testMergeDownOnLastColumnResetOncePreviousBlockCompleted() {
        let row1 = ExcelRow([ExcelCell("cell1 row1"),ExcelCell("cell2 row1 mergedown 2", [], rowspan: 2)])
        let row2 = ExcelRow([ExcelCell("cell1 row2")])
        let row3 = ExcelRow([ExcelCell("cell1 row3")])
        let row4 = ExcelRow([ExcelCell("cell1 row4"),ExcelCell("cell2 row4 mergedown 2", [], rowspan: 2)])
        let sheet1 = ExcelSheet([row1,row2,row3,row4], name: "Sheet1")
        
        let (exportResultCalled, url) = export([sheet1])
        
        XCTAssertTrue(exportResultCalled, "No file created.")
        XCTAssertEqual(valueOn(url, row: 4, cell: 2), nil)
    }
    
    func testAdjacentMergeDownOfDifferentSize() {
        let row1 = ExcelRow([ExcelCell("cell1 row1 mergedown 1", [], rowspan: 1),ExcelCell("cell2 row1 mergedown 2", [], rowspan: 2),ExcelCell("cell3 row1")])
        let row2 = ExcelRow([ExcelCell("cell1 row2 (idx 3)"), ExcelCell("cell2 row2 (idx4)")])
        let row3 = ExcelRow([ExcelCell("cell1 row3 (idx nil)"), ExcelCell("cell2 row3 (idx3)")])
        let row4 = ExcelRow([ExcelCell("cell1 row4 (idx nil)"), ExcelCell("cell2 row4 (nil)")])
        let sheet = ExcelSheet([row1,row2,row3,row4], name: "Sheet1")
        
        let (_, url) = export([sheet])
        
        XCTAssertEqual(valueOn(url, row: 2, cell: 1), "3")
        XCTAssertEqual(valueOn(url, row: 2, cell: 2), "4")
        XCTAssertEqual(valueOn(url, row: 3, cell: 1), nil)
        XCTAssertEqual(valueOn(url, row: 3, cell: 2), "3")
        XCTAssertEqual(valueOn(url, row: 4, cell: 1), nil)
        XCTAssertEqual(valueOn(url, row: 4, cell: 2), nil)
    }

    func testExport() {
        // arrange
        let cells = [ExcelCell("Age : "), ExcelCell("50", [TextAttribute.backgroundColor(Color.yellow), TextAttribute.font([TextAttribute.FontStyle.bold])])]
        let sheet = ExcelSheet([ExcelRow(cells), ExcelRow(cells), ExcelRow(cells)], name: "Test")
        let sheet2 = ExcelSheet([ExcelRow(cells)], name: "Test2")
        
        // act
        let (exportResultCalled, _) = export([sheet, sheet2], filename: "testExport")
        
        // assert
        XCTAssertTrue(exportResultCalled, "The file has not been created.")
    }
    
    class ParserDelegate : NSObject, XMLParserDelegate {
        var sheetNumberToCheck = 0
        var rowNumberToCheck = 0
        var cellNumberToCheck = 0
        var attributeName = "ss:Index"
        private var currentSheet = 0
        private var currentRow = 0
        private var currentCell = 0
        var valueFound : String? = nil
        
        func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?, qualifiedName qName: String?, attributes attributeDict: [String : String]) {
            switch elementName {
                case "Worksheet":
                    currentSheet+=1
                    currentRow = 0
                    currentCell = 0
                case "Row":
                    currentRow+=1
                    currentCell=0
                case "Cell":
                    currentCell+=1
                    if currentSheet == sheetNumberToCheck && currentRow == rowNumberToCheck && currentCell == cellNumberToCheck {
                        valueFound = attributeDict[attributeName]
                    }
                default: break
            }
        }
        
        func parser(_ parser: XMLParser, foundCharacters string: String) {
            
        }
        
        func parser(_ parser: XMLParser, didEndElement elementName: String, namespaceURI: String?, qualifiedName qName: String?) {
            
        }
        
        func parserDidEndDocument(_ parser: XMLParser) {
            
        }
    }

    fileprivate func valueOn(_ url: URL?, sheet: Int = 1, row: Int, cell: Int, forAttribute: String = "ss:Index") -> String? {
        do {
            let xmlData = try Data(contentsOf: url!)
            let parser = XMLParser(data: xmlData)
            let parserDelegate = ParserDelegate()
            parserDelegate.sheetNumberToCheck = sheet
            parserDelegate.rowNumberToCheck = row
            parserDelegate.cellNumberToCheck = cell
            
            parser.delegate = parserDelegate
            
            parser.parse()
            
            return parserDelegate.valueFound
        } catch let error {
            print("Error: \(error)")
        }
        
        return nil
    }
    
    fileprivate func export(_ sheets : [ExcelSheet], filename : String = #function) -> (Bool,URL?)
    {
        let expectation = self.expectation(description: "Async creation of Excel file.")
        var exportResultCalled: Bool = false
        var contentUrl: URL? = nil
        ExcelExport.export(sheets, fileName: filename) { url in
            print("done function called with \(String(describing: url))")
            if let file = url {
                do {
                    self.attachResult(content: file, name: filename)
                    contentUrl = file
                    exportResultCalled = try file.checkResourceIsReachable()
                    print("... the file \(file) exist and is reachable \(exportResultCalled).")
                }
                catch {
                    print("... url not reachable (not a file on the local system).")
                }
            }
            else {
                print("... no file produced.")
            }
            expectation.fulfill()
        }
        waitForExpectations(timeout: 1)
        
        return (exportResultCalled,contentUrl)
    }
    
    fileprivate func attachResult(content: URL, name: String)
    {
        let attachment = XCTAttachment(contentsOfFile: content)
        attachment.name = name
        attachment.lifetime = .keepAlways
        self.add(attachment)
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
