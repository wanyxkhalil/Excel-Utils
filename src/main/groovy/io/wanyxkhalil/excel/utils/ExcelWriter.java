package io.wanyxkhalil.excel.utils;

import io.wanyxkhalil.excel.utils.writer.WriterBuilder;

public class ExcelWriter {

    public static WriterBuilder builder() {
        return new WriterBuilder();
    }
}
