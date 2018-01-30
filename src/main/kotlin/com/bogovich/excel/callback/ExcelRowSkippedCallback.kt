package com.bogovich.excel.callback

interface ExcelRowSkippedCallback {
    fun skip(rowNum: Int)
}