package com.retheviper.excel.test.header

import com.retheviper.excel.definition.annotation.Column
import com.retheviper.excel.definition.annotation.Sheet
import java.util.*

@Sheet(dataStartIndex = 2)
data class Contract(

    @Column(position = "B")
    val name: String,

    @Column(position = "C")
    val companyPhone: String,

    @Column(position = "D")
    val cellPhone: String,

    @Column(position = "E")
    val homePhone: String,

    @Column(position = "F")
    val email: String,

    @Column(position = "G")
    val birth: Date,

    @Column(position = "H")
    val address: String,

    @Column(position = "I")
    val city: String,

    @Column(position = "J")
    val province: String,

    @Column(position = "K")
    val post: Int,

    @Column(position = "L")
    val memo: String? = null
)