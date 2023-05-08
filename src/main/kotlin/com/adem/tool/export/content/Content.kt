package com.adem.tool.export.content

import com.adem.tool.export.removeLast
import com.adem.tool.export.toBase64

class Content(val filename: String, val type: String, val data: ContentData) {
    init {
        val fileExt = fileExt()
        if (MimeTypeMap.mimeTypeFromExtension(fileExt) != type)
            throw IllegalArgumentException("File extension[$fileExt] does not match with content type[$type]")
    }

    fun withFilename(fileName: String) = Content(fileName, type, data)
    fun withType(type: String) = Content(filename, type, data)

    fun fileResource() = filename.replace(fileExt(), "").removeLast()

    fun fileExt(): String {
        val parts = filename.split(".")
        if (parts.isEmpty()) throw IllegalArgumentException("File extension does not exist on file name[$filename]")
        return parts.last()
    }

    fun dataAsBase64() = data.byte!!.toBase64()

    override fun toString() = "Content(filename='$filename', type='$type')"
}
