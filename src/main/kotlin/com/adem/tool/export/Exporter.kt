package com.adem.tool.export

import com.adem.tool.export.content.Content

interface Exporter {
    fun export(data: Any): Content
}
