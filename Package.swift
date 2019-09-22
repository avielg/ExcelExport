// swift-tools-version:5.0
import PackageDescription

let package = Package(
    name: "ExcelExport",
    products: [
    .library(name: "ExcelExport", targets: ["ExcelExport"])
        ],
    dependencies: [],
    targets: [
    .target(name: "ExcelExport"),
    .testTarget(name: "ExcelExportTests",
                dependencies: ["ExcelExport"])
    ]
)
