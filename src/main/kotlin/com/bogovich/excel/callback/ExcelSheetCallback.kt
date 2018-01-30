package com.bogovich.excel.callback

interface ExcelSheetCallback {
    fun startSheet(sheetIndex: Int, sheetName:String)
    fun endSheet(sheetIndex: Int, sheetName:String)
}