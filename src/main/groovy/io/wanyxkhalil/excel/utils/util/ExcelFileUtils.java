package io.wanyxkhalil.excel.utils.util;

import io.wanyxkhalil.excel.utils.exception.EuException;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class ExcelFileUtils {

    private static final String DEFAULT_FILE_DIR = "/tmp";

    private static final String DEFAULT_FILE_NAME_SEPARATOR = "-";

    private static final String SUFFIX_XLSX = ".xlsx";

    /**
     * 目录错误
     */
    private static final String FILE_IS_NOT_DIRECTORY = "目录错误";

    public static Path getXlsxPath(String name) {
        Path dirPath = Paths.get(ExcelFileUtils.DEFAULT_FILE_DIR);

        if (!Files.isDirectory(dirPath)) {
            throw EuException.build(FILE_IS_NOT_DIRECTORY);
        }

        StringBuilder sb = new StringBuilder();

        // 父路径
        sb.append(dirPath);
        sb.append(File.separator);

        // 文件名
        if (StringUtils.isNotBlank(name)) {
            sb.append(name);
            sb.append(DEFAULT_FILE_NAME_SEPARATOR);
        }
        sb.append(ExcelIDUtil.getInstance().nextId());

        // 文件后缀
        sb.append(ExcelFileUtils.SUFFIX_XLSX);

        return Paths.get(sb.toString());
    }
}
