package io.wanyxkhalil.excel.utils;

import io.wanyxkhalil.excel.utils.dynamicwriter.DynamicWriterBuilder;

public class ExcelDynamicWriter {

    public static DynamicWriterBuilder builder() {
        return new DynamicWriterBuilder();
    }
}
