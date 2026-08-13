import Foundation
import PDFKit
import AppKit

let args = CommandLine.arguments
guard args.count == 3 else { fatalError("usage: render_pdf_pages.swift input.pdf output_dir") }
let input = URL(fileURLWithPath: args[1])
let output = URL(fileURLWithPath: args[2], isDirectory: true)
try FileManager.default.createDirectory(at: output, withIntermediateDirectories: true)
guard let doc = PDFDocument(url: input) else { fatalError("cannot open pdf") }
for i in 0..<doc.pageCount {
    guard let page = doc.page(at: i) else { continue }
    let box = page.bounds(for: .mediaBox)
    let width: CGFloat = 1200
    let height = width * box.height / box.width
    let image = page.thumbnail(of: NSSize(width: width, height: height), for: .mediaBox)
    guard let tiff = image.tiffRepresentation,
          let rep = NSBitmapImageRep(data: tiff),
          let png = rep.representation(using: .png, properties: [:]) else { continue }
    let name = String(format: "page-%02d.png", i + 1)
    try png.write(to: output.appendingPathComponent(name))
}
print("pages=\(doc.pageCount)")
