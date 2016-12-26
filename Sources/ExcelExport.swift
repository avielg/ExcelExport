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


#if os(OSX)
    import AppKit
    public typealias Color = NSColor
#else
    import UIKit
    public typealias Color = UIColor
#endif



public enum TextAttribute {
    
    public enum FontStyle {
        case color(Color), bold
        
        var parsed: String {
            switch self {
            case .bold: return "ss:Bold=\"1\""
            case .color(let c): return "ss:Color=\"\(c.hexString())\""
            }
        }
    }
    
    case backgroundColor(Color)
    case font([FontStyle])
    case format(String)
    
    var parsed: String {
        switch self {
        case .backgroundColor(let c):
            return "<Interior ss:Color=\"\(c.hexString())\" ss:Pattern=\"Solid\"/>"
        
        case .format(let s):
            return "<NumberFormat ss:Format=\"\(s)\"/>"
            
        case .font(let styles):
            return "<Font " + styles.map{$0.parsed}.joined(separator: " ") + "/>"
        }
    }
    
    static func styleValue(for textAttributes: [TextAttribute]) -> String {
        guard textAttributes.count > 0 else { return "" }
        
        let parsedAttributes = textAttributes.map{ $0.parsed }
        return parsedAttributes.joined()
    }
}

extension TextAttribute {
    static let dateTimeTypeDateFormat = "1899-12-31T15:31:00.000"
}


public struct ExcelCell {
    let value: String
    let attributes: [TextAttribute]
    let colspan: Int?
    
    public enum DataType: String { case string="String", dateTime="DateTime" }
    let type: DataType
    
    public init(_ value: String) {
        self.value = value
        attributes = []
        colspan = nil
        type = .string
    }
    
    public init(_ value: String, _ attributes: [TextAttribute], _ type: DataType = .string, colspan: Int? = nil) {
        self.value = value
        self.attributes = attributes
        self.colspan = colspan
        self.type = type
    }
}


public struct ExcelRow {
    let cells: [ExcelCell]
    let height: Int?
    
    public init(_ cells: [ExcelCell], height: Int? = nil) {
        self.cells = cells
        self.height = height
    }
}

public struct ExcelSheet {
    
    enum Options {
        case printGridlines
        
        var parsed: String {
            switch self {
            case .printGridlines: return Parser.x.print(Parser.x.gridLines())
            }
        }
    }
    
    let rows: [ExcelRow]
    let name: String
    let options: [Options]
    
    public init(_ rows: [ExcelRow], name: String) {
        self.rows = rows
        self.name = name
        self.options = []
    }
}


public class ExcelExport {
    
    public class func export(_ sheets: [ExcelSheet], fileName: String, done: @escaping (URL?)->Void) {
        DispatchQueue.global(qos: .background).async {
            done(performXMLExport(sheets, fileName: fileName))
        }
    }
    
    private class func performXMLExport(_ sheets: [ExcelSheet], fileName: String) -> URL? {
        let file = fileUrl(name: fileName)
        
        // all styles for this wokrbook
        var styles = [String]()
        
        // adds new style, returns it's ID
        let appendStyle: (String)->String = {
            let id = "s\(styles.count)"
            styles.append("<Style ss:ID=\"\(id)\">\($0)</Style>")
            return id
        }
        
        var sheetsValues = [String]()
        for sheet in sheets {
            
            // build sheet
            var rows = [String]()
            for row in sheet.rows {

                var cells = [String]()
                for cell in row.cells {
                    
                    //data
                    let data = "<Data ss:Type=\"\(cell.type.rawValue)\">\(cell.value)</Data>"
                    
                    //style
                    let styleValue = TextAttribute.styleValue(for: cell.attributes)
                    let styleId = appendStyle(styleValue)
                    
                    let merge = cell.colspan.map{ "ss:MergeAcross=\"\($0)\"" } ?? ""
                    
                    //combine
                    let lead = "<Cell ss:StyleID=\"\(styleId)\" \(merge)>"
                    let trail = "</Cell>"
                    
                    cells.append([lead, data, trail].joined())
                }
                
                
                let rowOps = row.height.map{ "ss:Height=\"\($0)\"" } ?? ""
                let lead = "<Row \(rowOps)>"
                let trail = "</Row>"
                rows.append([lead, cells.joined(), trail].joined())
            }
            
            // combine
            let lead = "<Worksheet ss:Name=\"\(sheet.name)\"><Table>"
            let trail = "</Table></Worksheet>"
            sheetsValues.append([lead, rows.joined(), trail].joined())
        }
        
        
        let workbookLead = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><?mso-application progid=\"Excel.Sheet\"?><Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">"
        let workbookTrail = "</Workbook>"
        
        let stylesValue = "<Styles>\(styles.joined())</Styles>"
        
        let content = [workbookLead, stylesValue, sheetsValues.joined(), workbookTrail].joined()
        let totalRows = sheets.flatMap{ $0.rows }.count
        
        // write content to file
        do {
            try content.write(to: file, atomically: true, encoding: .utf8)
            print("\(totalRows) Lines written to file")
            return file
        } catch {
            print("Can't write \(totalRows) to file! [\(error)]")
            return nil
        }
    }
    
    private class func performExport(_ sheet: ExcelSheet, fileName: String) -> URL? {
        let file = fileUrl(name: fileName)
        
        // build rows
        let rowsValue = sheet.rows.map(Parser.build).joined()
        
        
        // build file content
        let templatePreLead = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'><head><meta http-equiv=Content-Type content='text/html; charset=UTF-8'><!--[if gte mso 9]><xml>"
        let templatePreTrail = "</xml><![endif]--></head><body><table>"
        
        let worksheetName = Parser.x.name(sheet.name)
        let worksheetOps = Parser.build(sheet.options)
        let worksheet = Parser.x.excelWorksheet(worksheetName + worksheetOps)
        let workbook = Parser.x.excelWorkbook(Parser.x.excelWorksheets(worksheet))
        
        let templatePre = [templatePreLead, workbook, templatePreTrail].joined()
        let templatePost = "</table border=\"1\"></body></html>"
        
        let content = [templatePre, rowsValue, templatePost].joined()
        
        // write content to file
        do {
            try content.write(to: file, atomically: true, encoding: .utf8)
            print("\(sheet.rows.count) Lines written to file")
            return file
        } catch {
            print("Can't write \(sheet.rows.count) to file! [\(error)]")
            return nil
        }
    }
    
    
    class func fileUrl(name: String) -> URL {
        let docsDir = FileManager.default.urls(for: .documentDirectory, in: .userDomainMask).first!
        return docsDir.appendingPathComponent("\(name).xls")
    }
    
}

class Parser {
    class func build(_ cell: ExcelCell) -> String {
        let cols = cell.colspan.map{ " colspan=\"\($0)\"" } ?? ""
        let style = TextAttribute.styleValue(for: cell.attributes)
        return "<td\(cols) \(style)>\(cell.value)</td>"
    }
    
    class func build(_ row: ExcelRow) -> String {
        return row.cells
            .map(Parser.build)
            .reduce("<tr>", +)
            .appending("</tr>")
    }
    
    class func build(_ sheetOptions: [ExcelSheet.Options]) -> String {
        let ops = sheetOptions
            .map(Parser.build)
            .joined()
        return x.worksheetOptions(ops)
    }
    
    class func build(_ sheetOption: ExcelSheet.Options) -> String {
        return sheetOption.parsed
    }
    
    /// parser for excel syntax 'x' tags
    struct x {
        
        // wrap
        static func name(_ v: String) -> String { return parse("Name", v) }
        static func print(_ v: String) -> String { return parse("Print", v) }
        static func worksheetOptions(_ v: String) -> String { return parse("WorksheetOptions", v) }
        static func excelWorksheet(_ v: String) -> String { return parse("ExcelWorksheet", v) }
        static func excelWorksheets(_ v: String) -> String { return parse("ExcelWorksheets", v) }
        static func excelWorkbook(_ v: String) -> String { return parse("ExcelWorkbook", v) }
        
        //single
        static func gridLines() -> String { return single("GridLines", "") }
        
        //general
        static func single(_ name: String, _ v: String) -> String {
            return "<x:\(name)/>"
        }
        static func parse(_ name: String, _ v: String) -> String {
            return "<x:\(name)>\(v)</x:\(name)>"
        }
    }

}

private extension Color {
    
    /// Hex string of a UIColor instance.
    ///
    /// - Parameter includeAlpha: Whether the alpha should be included.
    /// - Returns: HEX, including the '#'
    func hexString(_ includeAlpha: Bool = false) -> String {
        var r: CGFloat = 0
        var g: CGFloat = 0
        var b: CGFloat = 0
        var a: CGFloat = 0
        self.getRed(&r, green: &g, blue: &b, alpha: &a)
        
        if (includeAlpha) {
            return String(format: "#%02X%02X%02X%02X", Int(r * 255), Int(g * 255), Int(b * 255), Int(a * 255))
        } else {
            return String(format: "#%02X%02X%02X", Int(r * 255), Int(g * 255), Int(b * 255))
        }
    }
    
}

