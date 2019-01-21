package io.wanyxkhalil.excel.utils.util;

import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;

public class FileUtils {

    /**
     * 默认文件夹
     */
    private static final String DEFAULT_FILE_DIR = File.separator + "tmp" + File.separator;


    public static Path retrievePath(String name) {
        String path = DEFAULT_FILE_DIR +
                name + "-" + SnowflakeIDUtil.getInstance().nextId() +
                ".xlsx";
        return Paths.get(path);
    }
}
