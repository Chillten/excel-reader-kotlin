package com.bogovich.excel.callback

interface ExcelSheetSkippedCallback {
    fun skip(sheetNum: Int, sheetName: String)
}