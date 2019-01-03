package io.wanyxkhalil.excel.utils.domain;

import io.wanyxkhalil.excel.utils.normal.ExcelUtil;

import java.nio.file.Path;
import java.util.List;

public class ExcelObject {

    private Path path;
    private List<SheetObject> sheets;

    public Path write2File() {
        return ExcelUtil.write2File(path, sheets);
    }

    protected Path getPath() {
        return path;
    }

    public void setPath(Path path) {
        this.path = path;
    }

    protected List<SheetObject> getSheets() {
        return sheets;
    }

    public void setSheets(List<SheetObject> sheets) {
        this.sheets = sheets;
    }
}
