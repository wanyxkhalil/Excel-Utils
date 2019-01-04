package io.wanyxkhalil.excel.utils.testdomain;

import io.wanyxkhalil.excel.utils.annotation.ExcelField;

import java.math.BigDecimal;

public class DomainJava {

    @ExcelField("字符串")
    private String field1;
    @ExcelField("Integer")
    private Integer field2;
    @ExcelField("Double")
    private Double field3;
    @ExcelField("BigDecimal")
    private BigDecimal field4;

    public String getField1() {
        return field1;
    }

    public void setField1(String field1) {
        this.field1 = field1;
    }

    public Integer getField2() {
        return field2;
    }

    public void setField2(Integer field2) {
        this.field2 = field2;
    }

    public Double getField3() {
        return field3;
    }

    public void setField3(Double field3) {
        this.field3 = field3;
    }

    public BigDecimal getField4() {
        return field4;
    }

    public void setField4(BigDecimal field4) {
        this.field4 = field4;
    }
}
