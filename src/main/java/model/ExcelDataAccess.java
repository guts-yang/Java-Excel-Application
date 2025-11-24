package model;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Excel数据访问类 - MVC架构中的Model层组件
 * 负责Excel文件的读取和写入操作，实现数据持久化
 */
public class ExcelDataAccess {

    /**
     * 从Excel文件读取数据
     * @param filePath Excel文件路径
     * @return 包含表头和数据的对象数组，第一个元素是表头列表，第二个元素是数据列表
     * @throws IOException 文件操作异常
     */
    public Object[] readExcel(String filePath) throws IOException {
        List<String> headers = new ArrayList<>();
        List<List<String>> data = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            // 获取第一个工作表
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                return new Object[]{headers, data};
            }

            Iterator<Row> rowIterator = sheet.iterator();
            boolean isFirstRow = true;

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                List<String> rowData = new ArrayList<>();

                // 获取当前行的最大列数
                int lastCellNum = row.getLastCellNum();
                
                for (int cellIndex = 0; cellIndex < lastCellNum; cellIndex++) {
                    Cell cell = row.getCell(cellIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String cellValue = getCellValueAsString(cell);
                    
                    if (isFirstRow) {
                        // 第一行作为表头
                        headers.add(cellValue);
                    } else {
                        rowData.add(cellValue);
                    }
                }

                if (!isFirstRow && !rowData.isEmpty()) {
                    data.add(rowData);
                }
                isFirstRow = false;
            }
        }

        return new Object[]{headers, data};
    }

    /**
     * 将数据写入Excel文件
     * @param filePath 目标文件路径
     * @param headers 表头数据
     * @param data 表格数据
     * @throws IOException 文件操作异常
     */
    public void writeExcel(String filePath, List<String> headers, List<List<String>> data) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(new File(filePath))) {

            // 创建工作表
            Sheet sheet = workbook.createSheet("数据");

            // 创建表头行
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers.get(i));
                
                // 设置表头样式
                CellStyle headerStyle = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setBold(true);
                headerStyle.setFont(font);
                headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
                headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                headerStyle.setBorderBottom(BorderStyle.THIN);
                headerStyle.setBorderTop(BorderStyle.THIN);
                headerStyle.setBorderLeft(BorderStyle.THIN);
                headerStyle.setBorderRight(BorderStyle.THIN);
                
                cell.setCellStyle(headerStyle);
            }

            // 写入数据行
            for (int rowIndex = 0; rowIndex < data.size(); rowIndex++) {
                Row dataRow = sheet.createRow(rowIndex + 1); // +1 因为第一行是表头
                List<String> rowData = data.get(rowIndex);
                
                for (int cellIndex = 0; cellIndex < rowData.size(); cellIndex++) {
                    Cell cell = dataRow.createCell(cellIndex);
                    cell.setCellValue(rowData.get(cellIndex));
                    
                    // 设置数据单元格样式
                    CellStyle dataStyle = workbook.createCellStyle();
                    dataStyle.setBorderBottom(BorderStyle.THIN);
                    dataStyle.setBorderTop(BorderStyle.THIN);
                    dataStyle.setBorderLeft(BorderStyle.THIN);
                    dataStyle.setBorderRight(BorderStyle.THIN);
                    
                    cell.setCellStyle(dataStyle);
                }
            }

            // 自动调整列宽
            for (int i = 0; i < headers.size(); i++) {
                sheet.autoSizeColumn(i);
            }

            // 写入文件
            workbook.write(fos);
        }
    }

    /**
     * 获取单元格的值作为字符串
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    // 处理数字，避免科学计数法
                    double value = cell.getNumericCellValue();
                    if (value == Math.floor(value)) {
                        return String.valueOf((long) value);
                    } else {
                        return String.valueOf(value);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return String.valueOf(cell.getNumericCellValue());
                } catch (IllegalStateException e) {
                    return cell.getStringCellValue();
                }
            default:
                return "";
        }
    }
}