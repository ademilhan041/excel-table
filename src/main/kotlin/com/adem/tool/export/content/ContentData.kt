package com.adem.tool.export.content

import com.adem.tool.export.isNotNull
import com.adem.tool.export.safeClose
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.io.InputStream

class ContentData(val stream: InputStream?, val byte: ByteArray?) {
    fun size() = byte!!.size

    fun dataHashCode() = byte!!.hashCode()

    fun close() = stream?.safeClose()

    fun isStreamExist() = stream.isNotNull()

    fun isByteArrayExist() = byte.isNotNull() && byte!!.isNotEmpty()

    companion object {
        fun stream(d: InputStream): ContentData {
            val os = ByteArrayOutputStream()
            d.copyTo(os)
            return ContentData(d, os.toByteArray())
        }

        fun byte(d: ByteArray) = ContentData(ByteArrayInputStream(d), d)
    }

    override fun toString(): String {
        return "ContentData(stream=${isStreamExist()}, byte=${isByteArrayExist()})"
    }
}
