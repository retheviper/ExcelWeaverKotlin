package com.retheviper.excel.definition

import com.retheviper.excel.definition.type.DataType

/**
 * Definition of one cell.
 */
data class CellDef(

    /**
     * Cell's name.
     */
    val name: String,

    /**
     * Cell's data type.
     */
    val type: DataType,

    /**
     * Cell's index. (Horizontal)
     */
    val cellIndex: Int
)