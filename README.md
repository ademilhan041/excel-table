# Excel Table

### What is Excel Table?

The Excel Table Library is a tool that allows you to export data from your application into an Excel file format. It
provides a simple and easy-to-use interface for generating Excel files, without the need for complex coding or manual
formatting.

### Why Excel Table?

Creating excel files in Java can be a challenging task, primarily due to the limited availability of comprehensive
documentation of excel libraries. The Excel Table Library is a tool that allows you to export data from your
application into an Excel file format. It provides a simple and easy-to-use interface for generating Excel files,
without the need for complex coding or manual formatting.

* **Easy to use**: The Excel Table Library is designed to be user-friendly and straightforward. You don't need to
  have any prior experience with Excel programming or learn any new Excel libraries to use it.

* **No additional libraries required**: Unlike other Excel libraries that require you to download and install additional
  libraries, the Excel Table Library is entirely self-contained. You can get started right away, without having to
  worry about dependencies or compatibility issues.

### How easy is Excel Table?

```
val xlsExporter = ExporterBuilder()
    .xls(XlsConfig(mapOf("export.xls.type" to "POI").toProperties()))
    .build()

val content = xlsExporter.export(
    XlsData(
        XlsSheet(
            "Example",
            XlsLayout(
                mergeList = listOf(
                    XlsMerge(1, 2, 2, 2),
                    XlsMerge(1, 2, 3, 4)
                ),
                widthAdjustment = XlsColumnWidth(
                    columnWidth = listOf(20, 5)
                )
            ),
            XlsRow(
                XlsCell(
                    "Example Text", XlsCellOptions(
                        alignment = ElementAlignment.LEFT,
                        background = Color.TAN,
                        cellFont = CellFont.CALIBRI,
                    )
                ),
                XlsEmptyCell(),
                XlsCell(
                    1, XlsCellOptions(
                        type = ElementType(Type.NUMBER),
                        alignment = ElementAlignment.RIGHT,
                        cellFont = CellFont.ARIAL,
                    )
                ),
            )
        )
    )
)
```