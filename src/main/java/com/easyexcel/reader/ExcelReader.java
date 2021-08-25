package com.easyexcel.reader;

import com.easyexcel.annotation.ExcelColumn;
import com.easyexcel.exception.FieldNotFoundException;
import com.easyexcel.util.ObjectUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;


public class ExcelReader<T> {

    public static final String XLS = "xls";
    public static final String XLSX = "xlsx";

    private final Class<T> clazz;
    private final String type;
    private final OutputStream os;

    public ExcelReader(Class<T> clazz, String type, OutputStream os) {
        this.clazz = clazz;
        this.type = type;
        this.os = os;
    }

    /**
     * 导出excel
     *
     * @param data 导出数据
     */
    public void write(List<T> data) throws IOException {
        writeCustom(data, null);
    }

    /**
     * 导出excel 文件, 行和列旋转
     *
     * @param data      导出数据
     * @param filedName 作为第一行的字段名称
     */
    public void writeRotate(List<T> data, String filedName) throws IOException {
        Workbook workbook = getWorkbookObject();
        Sheet sheet = workbook.createSheet();
        if (ObjectUtils.isEmpty(data)) {
            return;
        }
        Field specifyFiled;
        try {
            specifyFiled = clazz.getDeclaredField(filedName);
        } catch (NoSuchFieldException e) {
            throw new FieldNotFoundException("字段名不存在, filed: " + filedName);
        }


        //  根据filedName, 创建标题行
        List<Row> rows = new ArrayList<>();
        // 第一个行应给为指定类的数据
        List<Field> fields = sortedFields(clazz);
        boolean afterSpecifyField = false;
        int specifyRowIdx = specifyFiled.equals(fields.get(0)) ? 0 : 1;

        // 标题列
        for (int i = 0; i < fields.size(); i++) {
            // 指定的字段在第一行, 按照注解上的排序, 在指定字段后面的-1, 在指定字段前面的+1
            Row row = sheet.createRow(i + (afterSpecifyField ? -specifyRowIdx : specifyRowIdx));
            if (filedName.equals(fields.get(i).getName())) {
                afterSpecifyField = true;
                row = sheet.createRow(0);
            }
            rows.add(row);
            Cell cell = row.createCell(0);
            cell.setCellValue(fields.get(i).getAnnotation(ExcelColumn.class).value());
        }

        // 数据列
        for (int i = 0; i < data.size(); i++) {
            for (int j = 0; j < fields.size(); j++) {
                Row row = rows.get(j);
                Cell cell = row.createCell(i + 1);
                Field field = fields.get(j);
                try {
                    boolean accessible = field.isAccessible();
                    field.setAccessible(true);
                    cell.setCellValue(Optional.ofNullable(field.get(data.get(i))).map(Object::toString).orElse(""));
                    field.setAccessible(accessible);
                } catch (IllegalAccessException ignored) {
                }
            }
        }


        workbook.write(os);
        os.flush();


    }

    public void writeCustom(List<T> data, List<String> customExportFields) throws IOException {
        Workbook workbook = getWorkbookObject();
        Sheet sheet = workbook.createSheet();
        if (ObjectUtils.isEmpty(data)) {
            return;
        }

        // 标题行
        Row titleRow = sheet.createRow(0);
        List<Field> fields = sortedFields(clazz);
        if(!ObjectUtils.isEmpty(customExportFields)) {
            fields = customExportFields.stream().map(e -> {
                try {
                    return clazz.getDeclaredField(e);
                } catch (NoSuchFieldException exception) {
                    throw new FieldNotFoundException("custom field not exist, please check! field name:{}" + e);
                }
            }).collect(Collectors.toList());
        }
        for (int i = 0; i < fields.size(); i++) {
            Field field = fields.get(i);
            if(field != null) {
                Cell cell = titleRow.createCell(i);
                cell.setCellValue(field.getAnnotation(ExcelColumn.class).value());
            }
        }

        for (int i = 0; i < data.size(); i++) {
            Row dataRow = sheet.createRow(i + 1);
            for (int j = 0; j < fields.size(); j++) {
                Field field = fields.get(j);
                Cell cell = dataRow.createCell(j);
                try {
                    boolean accessible = field.isAccessible();
                    field.setAccessible(true);
                    cell.setCellValue(Optional.ofNullable(field.get(data.get(i))).map(Object::toString).orElse(""));
                    field.setAccessible(accessible);
                } catch (IllegalAccessException ignored) {
                }
            }
        }

        workbook.write(os);
        os.flush();
    }

    private Workbook getWorkbookObject() {
        if (type.equalsIgnoreCase(XLS)) {
            return new HSSFWorkbook();
        }
        if (type.equalsIgnoreCase(XLSX)) {
            return new XSSFWorkbook();
        }
        throw new RuntimeException("文件格式不支持");
    }

    private List<Field> sortedFields(Class<T> clazz) {
        Class tempClazz = clazz;
        LinkedList<Field> list = new LinkedList<>();
        while (tempClazz != null) {
            list.addAll(0, Arrays.asList(tempClazz.getDeclaredFields()));
            tempClazz = tempClazz.getSuperclass();
        }

        return list.stream()
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                .sorted(Comparator.comparing(field -> field.getAnnotation(ExcelColumn.class).columnIdx()))
                .collect(Collectors.toList());
    }

}
