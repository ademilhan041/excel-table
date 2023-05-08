package com.adem.tool.export.xls

import com.adem.tool.export.Exporter
import com.adem.tool.export.content.Content
import com.adem.tool.export.xls.poi.XlsPoiExporter

class XlsExporter(private val xlsPoiExporter: XlsPoiExporter) : Exporter {
    override fun export(data: Any): Content {
        return xlsPoiExporter.export(data)
    }
}
