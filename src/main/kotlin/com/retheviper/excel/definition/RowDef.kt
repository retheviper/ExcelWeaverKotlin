package com.retheviper.excel.definition

/**
 * Definition of Row, which means 'Header'.
 */
data class RowDef(

    /**
     * Collection of each header cells.
     */
    val cellDefs: List<CellDef>
)