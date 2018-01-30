package com.bogovich.excel

import com.bogovich.excel.callback.ExcelReadCallback
import com.bogovich.excel.callback.ExcelSheetCallback
import com.bogovich.excel.callback.ExcelSheetSkippedCallback
import com.bogovich.excel.utils.OPCPackageUtils
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable
import org.apache.poi.xssf.eventusermodel.XSSFReader
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler
import org.apache.poi.xssf.model.StylesTable
import org.xml.sax.InputSource
import java.io.File
import java.io.InputStream
import javax.xml.parsers.SAXParserFactory

class ExcelReader(val xlsxPackage: OPCPackage, val sheetContentsHandler: XSSFSheetXMLHandler.SheetContentsHandler) {

    constructor(filePath: String, sheetContentsHandler: XSSFSheetXMLHandler.SheetContentsHandler) : this(OPCPackageUtils.getOPCPackage(filePath), sheetContentsHandler)
    constructor(inputStream: InputStream, sheetContentsHandler: XSSFSheetXMLHandler.SheetContentsHandler) : this(OPCPackageUtils.getOPCPackage(inputStream), sheetContentsHandler)
    constructor(file: File, sheetContentsHandler: XSSFSheetXMLHandler.SheetContentsHandler) : this(OPCPackageUtils.getOPCPackage(file), sheetContentsHandler)

    val executionContext = ExecutionContext()
    var sheetCallback: ExcelSheetCallback? = null
    var readCallback: ExcelReadCallback? = null
    var sheetSkippedCallback: ExcelSheetSkippedCallback? = null
    var sheetsToSkip: MutableList<String> = mutableListOf()


    private fun read(sheetNumber: Int) {
        listOfNotNull(sheetCallback, readCallback, sheetSkippedCallback, sheetContentsHandler)
                .forEach { callback ->
                    if (callback is ExecutionContextAware) {
                        callback.executionContext = this.executionContext
                    }
                }

        val strings = ReadOnlySharedStringsTable(this.xlsxPackage)
        val xssfReader = XSSFReader(this.xlsxPackage)
        val worksheets = xssfReader.sheetsData as XSSFReader.SheetIterator


        readCallback?.beforeReading()

        worksheets.withIndex().forEach { sheetInput ->
            if (worksheets.sheetName in sheetsToSkip) {
                sheetSkippedCallback?.skip(sheetInput.index, worksheets.sheetName)
            } else {
                this.sheetCallback?.startSheet(sheetInput.index, worksheets.sheetName)

                if (sheetNumber == -1 || sheetNumber == sheetInput.index) {
                    readSheet(xssfReader.stylesTable, strings, sheetInput.value)
                }

                this.sheetCallback?.endSheet(sheetInput.index, worksheets.sheetName)
            }
        }

    }

    private fun readSheet(stylesTable: StylesTable, sharedStringsTable: ReadOnlySharedStringsTable, inputStream: InputStream) {
        inputStream.use { stream ->
            val sheetReader = SAXParserFactory.newInstance().newSAXParser().xmlReader
            sheetReader.contentHandler = XSSFSheetXMLHandler(stylesTable, sharedStringsTable, sheetContentsHandler, true)
            sheetReader.parse(InputSource(stream))
        }
    }
}