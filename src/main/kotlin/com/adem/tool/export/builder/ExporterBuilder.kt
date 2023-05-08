package com.adem.tool.export.builder

import com.adem.tool.export.Exporter
import com.adem.tool.export.xls.*
import com.adem.tool.export.xls.poi.XlsPoiExporter

class ExporterBuilder {
    private lateinit var _xlsConfig: XlsConfig

    fun xls(config: XlsConfig): ExporterBuilder {
        this._xlsConfig = config
        return this
    }

    fun build(): Exporter {
        val xlsExporterImplementation = when (_xlsConfig.type) {
            XlsType.POI -> XlsPoiExporter()
        }

        return XlsExporter(xlsExporterImplementation)
    }
}

fun main() {
    ExporterBuilder().build().export(
        XlsData(
            XlsSheet(
                "Example",
                XlsLayout(
                    listOf(
                        XlsMerge(1, 2, 2, 2),
                        XlsMerge(1, 2, 3, 4)
                    )
                ),
                XlsRow(
                    XlsCell(
                        "Example Text", XlsCellOptions(
                            alignment = ElementAlignment.RIGHT,
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
}
