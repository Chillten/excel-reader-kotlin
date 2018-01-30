package com.bogovich.excel

class ExecutionContext {
    private val coreStore: Map<String, Any> = hashMapOf()
    val store: Map<String, Any> = hashMapOf()
    val errors: MutableList<ExcelFileReaderError> = mutableListOf()
    var iteration: Int = 0
    val chunkItems: MutableList<Any> = mutableListOf()
    val unprocessableSheets: MutableSet<Int> = mutableSetOf()
    var sheetIndex: Int = 0
    var sheetName: String = ""


    fun clean() {
        errors.clear()
        chunkItems.clear()
        unprocessableSheets.clear()
        iteration = 0
    }

    fun closeChunk() {
        iteration = 0
        chunkItems.clear()
    }
}