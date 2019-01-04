package io.wanyxkhalil.excel.utils.dynamic;

import io.wanyxkhalil.excel.utils.normal.ExcelBuilder;
import io.wanyxkhalil.excel.utils.util.ExcelFileUtils;

import java.nio.file.Path;

public class DynamicExcelBuilder extends ExcelBuilder {

    @Override
    public DynamicExcelObject build() {
        DynamicExcelObject obj = new DynamicExcelObject();
        obj.setSheets(this.sheets);

        Path path = ExcelFileUtils.getXlsxPath(this.fileName);
        obj.setPath(path);

        return obj;
    }
}