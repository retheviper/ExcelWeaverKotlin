package com.retheviper.excel.definition

import kotlin.reflect.KClass

/**
 * Definition of sheet.
 */
data class SheetDef(

    /**
     * Definition of row, which means header.
     */
    val rowDef: RowDef,

    /**
     * Sheet's name.
     */
    val name: String,

    /**
     * Sheet's index.
     */
    val index: Int,

    /**
     * Row index which data part begins.
     */
    val dataStartIndex: Int,

    /**
     * Object to map with this Sheet. The object will be considered as sheet, and fields are considered as cells.
     */
    val dataClass: Class<*>
)