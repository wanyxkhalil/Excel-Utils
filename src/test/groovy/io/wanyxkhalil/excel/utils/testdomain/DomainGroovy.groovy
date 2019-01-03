package io.wanyxkhalil.excel.utils.testdomain

import io.wanyxkhalil.excel.utils.config.ExcelField

class DomainGroovy {

    @ExcelField("字符串")
    String field1
    @ExcelField("Integer")
    Integer field2
    @ExcelField("Double")
    Double field3
    @ExcelField("BigDecimal")
    BigDecimal field4
}
