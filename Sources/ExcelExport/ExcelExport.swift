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

    public enum FontStyle: Equatable {
        case color(Color), bold

        var parsed: String {
            switch self {
            case .bold: return "ss:Bold=\"1\""
            case .color(let color): return "ss:Color=\"\(color.hexString())\""
            }
        }

        public static func == (lhs: FontStyle, rhs: FontStyle) -> Bool {
            switch (lhs, rhs) {
            case (.bold, .bold): return true
            case (.color(let lColor), .color(let rColor)): return lColor == rColor
            default: return false
            }
        }
    }

    case backgroundColor(Color)
    case font([FontStyle])
    case format(String)

    var parsed: String {
        switch self {
        case .backgroundColor(let color):
            return "<Interior ss:Color=\"\(color.hexString())\" ss:Pattern=\"Solid\"/>"

        case .format(let format):
            return "<NumberFormat ss:Format=\"\(format)\"/>"

        case .font(let styles):
            return "<Font " + styles.map({$0.parsed}).joined(separator: " ") + "/>"
        }
    }

    public static func a(lhs: TextAttribute, rhs: TextAttribute) -> Bool {
        switch (lhs, rhs) {
        case (.backgroundColor(let lColor), .backgroundColor(let rColor)): return lColor == rColor
        case (.font(let lFont), .font(let rFont)): return lFont == rFont
        case (.format(let lFormat), .format(let rFormat)): return lFormat == rFormat
        default: return false
        }
    }

    static func styleValue(for textAttributes: [TextAttribute]) -> String {
        guard textAttributes.count > 0 else { return "" }

        let parsedAttributes = textAttributes.map { $0.parsed }
        return parsedAttributes.joined()
    }
}

extension TextAttribute {
    static let dateTimeTypeDateFormat = "1899-12-31T15:31:00.000"
}

public struct ExcelCell {
    public let value: String
    public let attributes: [TextAttribute]
    public let colspan: Int?
    public let rowspan: Int?

    /**
     This date formatter is used to format the date in the Data element for a DateTime cell.
     It match what Excel is expecting.
     */
    public static let dateFormatter: DateFormatter = {
        let formatter = DateFormatter()
        formatter.locale = Locale(identifier: "en_US_POSIX")
        formatter.dateFormat = "yyyy-MM-dd'T'HH:mm:ss.SSS"
        return formatter
    }()

    public enum DataType: String { case string="String", dateTime="DateTime", number="Number" }
    let type: DataType

    public init(_ value: String, _ attributes: [TextAttribute], _ type: DataType = .string, colspan: Int? = nil,
                rowspan: Int? = nil) {
        self.value = value
        self.attributes = attributes
        self.colspan = colspan
        self.type = type
        self.rowspan = rowspan
    }

    public init(_ value: Double, _ attributes: [TextAttribute] = [], colspan: Int? = nil, rowspan: Int? = nil) {
        self.init(String(format: "%f", arguments: [value]), attributes, .number, colspan: colspan, rowspan: rowspan)
    }

    /**
     - Warnings: there is a default format (in attributes) as General Date.
     If you want to specify other attributes, add the desired date format too.
     */
    public init(_ value: Date, _ attributes: [TextAttribute] = [.format("General Date")], colspan: Int? = nil,
                rowspan: Int? = nil) {
        self.init( ExcelCell.dateFormatter.string(from: value), attributes, .dateTime, colspan: colspan,
                   rowspan: rowspan)
    }

    public init(_ value: String, _ attributes: [TextAttribute] = [], colspan: Int? = nil, rowspan: Int? = nil) {
        self.init( value, attributes, .string, colspan: colspan, rowspan: rowspan)
    }

    public var dataElement: String {
        return "<Data ss:Type=\"\(self.type.rawValue)\">\(self.value)</Data>"
    }

    public func cellElement(withStyleId styleId: String?, withIndexAttribute indexAttribute: String) -> String {
        let mergeAcross = self.colspan.map { " ss:MergeAcross=\"\($0)\"" } ?? ""
        let mergeDown = self.rowspan.map { " ss:MergeDown=\"\($0)\"" } ?? ""
        let style = styleId != nil ? " ss:StyleID=\"\(styleId!)\"" : ""
        let lead = "<Cell\(style)\(mergeAcross)\(mergeDown)\(indexAttribute)>"
        let trail = "</Cell>"

        return [lead, self.dataElement, trail].joined()
    }
}

public struct ExcelRow {
    public let cells: [ExcelCell]
    public let height: Int?

    public init(_ cells: [ExcelCell], height: Int? = nil) {
        self.cells = cells
        self.height = height
    }

    public var lead: String {
        let rowOps = self.height.map { "ss:Height=\"\($0)\"" } ?? ""
        let lead = "<Row \(rowOps)>"
        return lead
    }

    public let trail: String = "</Row>"
}

public struct ExcelSheet {
    public let rows: [ExcelRow]
    public let name: String

    public init(_ rows: [ExcelRow], name: String) {
        self.rows = rows
        self.name = name
    }

    public var lead: String {
        return "<Worksheet ss:Name=\"\(self.name)\"><Table>"
    }

    public let trail = "</Table></Worksheet>"
}

public class ExcelExport {
    private let workbookLead = """
                           <?xml version=\"1.0\" encoding=\"UTF-8\"?>\
                           <?mso-application progid=\"Excel.Sheet\"?>\
                           <Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" \
                           xmlns:x=\"urn:schemas-microsoft-com:office:excel\" \
                           xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" \
                           xmlns:html=\"http://www.w3.org/TR/REC-html40\">
                           """
    private let workbookTrail = "</Workbook>"

    private init() {

    }

    public class func export(_ sheets: [ExcelSheet], fileName: String, done: @escaping (URL?) -> Void) {
        DispatchQueue.global(qos: .background).async {
            let resultUrl = ExcelExport().performXMLExport(sheets, fileName: fileName)
            DispatchQueue.main.async { done(resultUrl) }
        }
    }

    private func performXMLExport(_ sheets: [ExcelSheet], fileName: String) -> URL? {
        var sheetsValues = [String]()
        for sheet in sheets {
            // build sheet
            var rows = [String]()
            for row in sheet.rows {
                var cells = [String]()
                startRow()
                for (cellIndex, cell) in row.cells.enumerated() {
                    computeCellIndex(cellIndex)

                    //style
                    let styleId = self.styleId(for: cell.attributes)

                    let indexAttribute = generateIndexAttribute(cellIndex)

                    cells.append(cell.cellElement(withStyleId: styleId, withIndexAttribute: indexAttribute))

                    setupMergeDownCells(cell)
                }
                decreaseRemainingRowSpanOnRemainingCells()

                rows.append([row.lead, cells.joined(), row.trail].joined())
            }

            // combine rows on sheet
            sheetsValues.append([sheet.lead, rows.joined(), sheet.trail].joined())

            remainingSpan = [RemainingSpan]()
        }

        return writeToFile(name: fileName, sheets: sheetsValues, totalRows: sheets.flatMap { $0.rows }.count)
    }

    // MARK: - Output functions
    private func writeToFile(name fileName: String, sheets sheetsValues: [String], totalRows: Int) -> URL? {
        let file = fileUrl(name: fileName)
        let content = [workbookLead, stylesValue, sheetsValues.joined(), workbookTrail].joined()

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

    private func fileUrl(name: String) -> URL {
        let docsDir = FileManager.default.urls(for: .documentDirectory, in: .userDomainMask).first!
        return docsDir.appendingPathComponent("\(name).xls")
    }

    // MARK: - Styles feature utility properties and functions

    // all styles for this workbook
    private var styles = [String: String]() // id : value

    // adds new style, returns it's ID
    private func appendStyle(_ styleValue: String) -> String {
        let id = "s\(styles.count)"
        styles[id] = "<Style ss:ID=\"\(id)\">\(styleValue)</Style>"
        return id
    }

    private func styleId(for attributes: [TextAttribute]) -> String? {
        let styleValue = TextAttribute.styleValue(for: attributes)
        let styleId: String?
        if styleValue.isEmpty {
            styleId = nil
        } else if let id = styles.first(where: { _, value in value.contains(styleValue) })?.key {
            styleId = id //reuse existing style
        } else {
            styleId = appendStyle(styleValue) //create new style
        }
        return styleId
    }

    private var stylesValue: String {
        return "<Styles>\(styles.values.joined())</Styles>"
    }

    // MARK: - MergeDown feature utility types, properties and functions
    struct RemainingSpan {
        var remainingRows: Int
        var colSpan: Int
        var description: String {
            return "remainingRows: \(remainingRows), colSpan: \(colSpan)"
        }
    }

    private var vIndex: Int = 0
    private var remainingSpan = [RemainingSpan]()

    private func startRow() {
        vIndex = 0
    }

    private func computeCellIndex(_ cellIndex: Int) {
        while vIndex < remainingSpan.count && remainingSpan[vIndex].remainingRows > 0 {
            remainingSpan[vIndex].remainingRows -= 1
            vIndex += (remainingSpan[vIndex].colSpan + 1)
        }
    }

    private func decreaseRemainingRowSpanOnRemainingCells() {
        while vIndex < remainingSpan.count {
            remainingSpan[vIndex].remainingRows -= 1
            vIndex += 1
        }
    }

    private func generateIndexAttribute(_ cellIndex: Int) -> String {
        return vIndex != cellIndex ? " ss:Index=\"\(vIndex+1)\"": ""
    }

    private func setupMergeDownCells(_ cell: ExcelCell) {
        // Setup mergeDown cells
        if let newMergeDownCount = cell.rowspan {
            while remainingSpan.count <= vIndex {
                remainingSpan.append(RemainingSpan(remainingRows: 0, colSpan: 0))
            }
            remainingSpan[vIndex] = RemainingSpan(remainingRows: newMergeDownCount,
                                                  colSpan: cell.colspan ?? 0)
        }
        vIndex += 1
    }
}

private extension Color {

    /// Hex string of a UIColor instance.
    ///
    /// - Parameter includeAlpha: Whether the alpha should be included.
    /// - Returns: HEX, including the '#'
    func hexString(_ includeAlpha: Bool = false) -> String {
        var red: CGFloat = 0
        var green: CGFloat = 0
        var blue: CGFloat = 0
        var alpha: CGFloat = 0
        self.getRed(&red, green: &green, blue: &blue, alpha: &alpha)

        if includeAlpha {
            return String(format: "#%02X%02X%02X%02X", Int(red * 255), Int(green * 255), Int(blue * 255),
                          Int(alpha * 255))
        } else {
            return String(format: "#%02X%02X%02X", Int(red * 255), Int(green * 255), Int(blue * 255))
        }
    }
}
