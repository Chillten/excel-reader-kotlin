package com.bogovich.excel.callback

interface ExcelRowContentCallback {
    fun processRow(rowNum: Int, map: Map<String, String>)
}