package com.retheviper.excel.definition.annotation


/**
 * Mark field as cell in sheet.
 */
@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.FIELD)
annotation class Column(

    /**
     * Cell's name. If null or "", it will be considered as field's name.
     */
    val name: String = "",

    /**
     * Cell's position, such as "A" or "AA". Necessary to input correct value.
     */
    val position: String = "A"
)