package com.retheviper.excel.definition.annotation


/**
 * Mark class as definition of Sheet.
 */
@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.CLASS)
annotation class Sheet(

    /**
     * Sheet's name. If null or "", it will be considered as class's name.
     */
    val name: String = "",

    /**
     * Sheet's index.
     */
    val index: Int = -1,

    /**
     * Row index which data part begins.
     */
    val dataStartIndex: Int = 0
)