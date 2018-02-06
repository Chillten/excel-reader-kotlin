package com.bogovich

import com.bogovich.excel.ExcelReader
import com.bogovich.excel.ExcelWorkSheetHandler
import org.junit.Test
import java.io.File

class TestSource {
    @Test fun f() {
        val file = File("C:/Users/aleksandr.bogovich/Desktop/my staff/practice/excel-poi-xssf-reader/src/test/resources/file/Sample-Person-Data.xlsx")

        val workSheetHandler = ExcelWorkSheetHandler()
        val reader = ExcelReader(file, workSheetHandler)
        reader.read()
    }
}