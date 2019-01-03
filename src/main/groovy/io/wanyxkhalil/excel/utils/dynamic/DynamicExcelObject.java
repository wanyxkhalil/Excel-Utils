package io.wanyxkhalil.excel.utils.dynamic;


import io.wanyxkhalil.excel.utils.domain.ExcelObject;

import java.nio.file.Path;

public class DynamicExcelObject extends ExcelObject {

    @Override
    public Path write2File() {
        return DynamicExcelUtil.write2File(getPath(), getSheets());
    }

}
