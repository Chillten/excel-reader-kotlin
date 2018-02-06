package com.bogovich.excel

import com.bogovich.excel.callback.ExcelRowContentCallback
import mu.KLogging

class ExcelWorkSheetHandler {
    companion object: KLogging()
    val head_row = 0;
    private var rowCallback: ExcelRowContentCallback? = null
    private var currentRow = 0


}