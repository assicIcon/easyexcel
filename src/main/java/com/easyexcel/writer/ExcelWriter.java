package com.easyexcel.writer;

import com.easyexcel.annotation.ExcelColumn;
import com.easyexcel.util.ObjectUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Type;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author wrh
 * @date 2021/8/25
 */
public class ExcelWriter<T> {

    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";

    private final Class<T> clazz;
    private final Map<String, Field> columnFieldMap;
    private Map<Integer, String> indexColumnNameMap;
    private final String fileName;
    private final InputStream is;

    public ExcelWriter(Class<T> clazz, String fileName, InputStream is) {
        this.clazz = clazz;
        this.columnFieldMap = columnFieldMap(clazz);
        this.fileName = fileName;
        this.is = is;
    }


    public List<T> read() {
        Workbook workbook = null;
        List<T> list = new ArrayList<>();
        try {
            // 获取workbook
            workbook = getWorkbookObject();
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                return Collections.emptyList();
            }

            // 获取第一行的列头, 及每列下标对应的列头名称
            int columnRowNum = sheet.getFirstRowNum();
            Row columnRow = sheet.getRow(columnRowNum);
            if (columnRow == null) {
                return Collections.emptyList();
            }
            this.indexColumnNameMap = indexColumnNameMap(columnRow);

            // 读取数据
            int dataRowStartNum = columnRowNum + 1;
            for (int i = dataRowStartNum; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    continue;
                }
                T object = getObjectFromRow(row);
                if (object != null) {
                    list.add(object);
                }
            }

        } catch (Exception e) {
//            log.error("read excel error, msg: ", e);
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return list;
    }

    private Workbook getWorkbookObject() throws IOException {
        String fileType = fileName.substring(fileName.lastIndexOf(".") + 1);
        if (fileType.equalsIgnoreCase(XLS)) {
            return new HSSFWorkbook(is);
        }
        if (fileType.equalsIgnoreCase(XLSX)) {
            return new XSSFWorkbook(is);
        }
        throw new RuntimeException("文件格式不支持");
    }

    private static <T> Map<String, Field> columnFieldMap(Class<T> clazz) {
        return Stream.of(clazz.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                .collect(Collectors.toMap(field -> field.getAnnotation(ExcelColumn.class).value().trim(), e -> e));
    }

    private static Map<Integer, String> indexColumnNameMap(Row columnRow) {
        Iterator<Cell> cellIterator = columnRow.cellIterator();
        Map<Integer, String> indexColumnNameMap = new HashMap<>();
        while (cellIterator.hasNext()) {
            Cell next = cellIterator.next();
            String cellValue = next.getStringCellValue();
            if (!ObjectUtils.isEmpty(cellValue)) {
                indexColumnNameMap.put(next.getColumnIndex(), cellValue);
            }
        }
        return indexColumnNameMap;
    }

    private T getObjectFromRow(Row row) throws InstantiationException, IllegalAccessException {
        T instance = null;
        for (Map.Entry<Integer, String> entry : indexColumnNameMap.entrySet()) {
            Integer idx = entry.getKey();
            String columnName = entry.getValue();
            Cell cell = row.getCell(idx);
            Field field = columnFieldMap.get(columnName.trim());
            if (field == null) {
                continue;
            }
            try {
                boolean accessible = field.isAccessible();
                field.setAccessible(true);
                Object cellValue = getCellValue(cell);
                if (cellValue == null) {
                    continue;
                }
                if (instance == null) {
                    instance = clazz.newInstance();
                }
                field.set(instance, toFieldRequireType(cellValue, field));
                field.setAccessible(accessible);
            } catch (Exception e) {
//                log.error("read excel row value filling to object fail, msg: ", e);
                e.printStackTrace();
            }
        }
        return instance;
    }

    private Object toFieldRequireType(Object value, Field field) {
        if (ObjectUtils.isEmpty(value)) {
            return null;
        }
        Type type = field.getGenericType();
        DataType dataType = DataType.of(type);
        try {
            return dataType.transferType(value);
        } catch (Exception e) {
//            log.error("data type transfer fail fieldType:{}, value:{}, valueType:{}", type, value, value.getClass(), e);
            e.printStackTrace();
        }
        return null;
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        CellType cellTypeEnum = cell.getCellTypeEnum();
        switch (cellTypeEnum) {
            case NUMERIC:
                return cell.getNumericCellValue();
            case STRING:
                return cell.getStringCellValue();
            case FORMULA:
                try {
                    return cell.getDateCellValue();
                } catch (Exception e) {
                    return cell.getStringCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case ERROR:
                return cell.getErrorCellValue();
            case _NONE:
            case BLANK:
            default:
                return null;
        }
    }


}

