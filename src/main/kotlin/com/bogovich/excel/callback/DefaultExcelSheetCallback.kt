package com.bogovich.excel.callback

import com.bogovich.excel.ExecutionContext
import com.bogovich.excel.ExecutionContextAware

class DefaultExcelSheetCallback : ExcelSheetCallback,
        ExecutionContextAware {

    override var executionContext: ExecutionContext? = null

    override fun startSheet(sheetIndex: Int, sheetName: String) {
        executionContext?.sheetIndex = sheetIndex
        executionContext?.sheetName = sheetName
    }

    override fun endSheet(sheetIndex: Int, sheetName: String) {
        executionContext?.sheetIndex = -1
        executionContext?.sheetName = ""
    }
}