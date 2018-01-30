package com.bogovich.excel

interface ExcelItemWriter<T> {
    fun writeItems(executionContext: ExecutionContext): Int
    fun writeLastItems(executionContext: ExecutionContext): Int
}