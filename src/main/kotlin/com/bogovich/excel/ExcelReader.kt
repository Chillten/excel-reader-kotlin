package com.bogovich.excel

import com.bogovich.excel.utils.OPCPackageUtils
import mu.KLogging
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable
import org.apache.poi.xssf.eventusermodel.XSSFReader
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler
import org.apache.poi.xssf.model.StylesTable
import org.xml.sax.InputSource
import java.io.File
import java.io.InputStream
import javax.xml.parsers.SAXParser
import javax.xml.parsers.SAXParserFactory

class ExcelReader(val xlsxPackage: OPCPackage, val sheetContentsHandler: XSSFSheetXMLHandler.SheetContentsHandler) {

    companion object : KLogging()

    constructor(filePath: String, sheetContentsHandler: XSSFSheetXMLHandler.SheetContentsHandler) : this(OPCPackageUtils.getOPCPackage(filePath), sheetContentsHandler)
    constructor(inputStream: InputStream, sheetContentsHandler: XSSFSheetXMLHandler.SheetContentsHandler) : this(OPCPackageUtils.getOPCPackage(inputStream), sheetContentsHandler)
    constructor(file: File, sheetContentsHandler: XSSFSheetXMLHandler.SheetContentsHandler) : this(OPCPackageUtils.getOPCPackage(file), sheetContentsHandler)


    fun read() {
        val strings = ReadOnlySharedStringsTable(this.xlsxPackage)
        val xssfReader = XSSFReader(this.xlsxPackage)
        val worksheets = xssfReader.sheetsData as XSSFReader.SheetIterator

        worksheets.forEach { inputStream ->
            inputStream.use {
                readSheet(xssfReader.stylesTable, strings, inputStream)
            }
        }
    }

    private fun readSheet(stylesTable: StylesTable, strings: ReadOnlySharedStringsTable, inputStream: InputStream) {
        val sheetsParser = SAXParserFactory.newInstance().newSAXParser().xmlReader
        sheetsParser.contentHandler = XSSFSheetXMLHandler(stylesTable, strings, sheetContentsHandler, true)
        sheetsParser.parse(InputSource(inputStream))
    }
}