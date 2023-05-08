package com.adem.tool

import com.adem.tool.export.builder.ExporterBuilder
import com.adem.tool.export.xls.*
import org.junit.Ignore
import org.junit.Test
import java.io.File

class XlsExporterTest {

    @Test
    @Ignore
    fun testXls() {
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

        val file = File("test.xlsx")
        file.writeBytes(content.data.byte!!)
    }
}