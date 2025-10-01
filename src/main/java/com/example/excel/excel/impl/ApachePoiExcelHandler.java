package com.example.excel.excel.impl;

import com.example.excel.excel.ExcelHandler;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;
import org.springframework.web.multipart.MultipartFile;

import cn.hutool.core.date.StopWatch;

import jakarta.servlet.http.HttpServletResponse;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Apache POI Excel处理器 - 简化版
 * 仅保留基本功能，用于特殊场景
 */
@Slf4j
public class ApachePoiExcelHandler<T> implements ExcelHandler<T> {

    @Override
    public void export(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) {
        try {
            // 确保文件名以.xlsx结尾
            if (!filename.toLowerCase().endsWith(".xlsx")) {
                filename = filename + ".xlsx";
            }
            
            // 设置响应头
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding("UTF-8");
            String encodedFileName = URLEncoder.encode(filename, StandardCharsets.UTF_8.toString());
            response.setHeader("Content-Disposition", "attachment; filename=" + encodedFileName);
            
            // 使用SXSSFWorkbook处理大数据量
            try (SXSSFWorkbook workbook = new SXSSFWorkbook(100)) {
                Sheet sheet = workbook.createSheet("Sheet1");
                
                // 获取类的所有字段
                Field[] fields = clazz.getDeclaredFields();
                
                // 创建表头
                Row headerRow = sheet.createRow(0);
                for (int i = 0; i < fields.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(fields[i].getName());
                }
                
                // 填充数据
                int rowIndex = 1;
                for (T data : dataList) {
                    Row row = sheet.createRow(rowIndex++);
                    for (int i = 0; i < fields.length; i++) {
                        fields[i].setAccessible(true);
                        Object value = fields[i].get(data);
                        Cell cell = row.createCell(i);
                        if (value != null) {
                            cell.setCellValue(value.toString());
                        }
                    }
                }
                
                // 写入响应
                try (OutputStream outputStream = response.getOutputStream()) {
                    workbook.write(outputStream);
                    outputStream.flush();
                }
            }
        } catch (Exception e) {
            log.error("Apache POI导出Excel失败", e);
            throw new RuntimeException("导出Excel失败", e);
        }
    }

    @Override
    public List<T> importExcel(MultipartFile file, Class<T> clazz) {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        
        try (InputStream inputStream = file.getInputStream()) {
            log.info("开始Apache POI导入Excel文件：{}，大小：{}KB", file.getOriginalFilename(), file.getSize() / 1024);
            
            // 检查文件大小
            long fileSize = inputStream.available();
            if (fileSize > 100_000_000) {
                log.warn("处理超大Excel文件(大小: {}MB)，可能需要较长时间", fileSize/1024/1024);
            }
            
            // 使用SAX处理器处理大文件
            ExcelSAXHandler<T> handler = new ExcelSAXHandler<>(clazz);
            List<T> result = handler.parse(inputStream);
            
            stopWatch.stop();
            log.info("Apache POI导入Excel完成，共导入{}条数据，耗时{}ms", result.size(), stopWatch.getTotalTimeMillis());
            
            return result;
        } catch (org.apache.poi.util.RecordFormatException e) {
            log.error("Excel文件格式异常，可能是文件损坏或超大文件：{}", e.getMessage());
            throw new RuntimeException("Excel文件格式异常，可能是文件损坏或超大文件，请尝试使用EasyExcel处理器或分割文件后重试", e);
        } catch (OutOfMemoryError e) {
            log.error("Excel导入过程中内存不足", e);
            throw new RuntimeException("Excel导入过程中内存不足，请尝试使用EasyExcel处理器或增加JVM内存", e);
        } catch (Exception e) {
            log.error("Apache POI导入Excel失败", e);
            throw new RuntimeException("导入Excel失败：" + e.getMessage(), e);
        } finally {
            // 强制回收内存
            System.gc();
        }
    }
    
    /**
     * 根据字段类型设置值
     */
    private void setFieldValue(Field field, T instance, Cell cell) {
        try {
            String cellValue = getCellValueAsString(cell);
            if (cellValue == null) {
                return;
            }

            Class<?> fieldType = field.getType();
            field.setAccessible(true);
            
            try {
                if (String.class.equals(fieldType)) {
                    field.set(instance, cellValue);
                } else if (Integer.class.equals(fieldType) || int.class.equals(fieldType)) {
                    field.set(instance, Integer.parseInt(cellValue));
                } else if (Long.class.equals(fieldType) || long.class.equals(fieldType)) {
                    field.set(instance, Long.parseLong(cellValue));
                } else if (Double.class.equals(fieldType) || double.class.equals(fieldType)) {
                    field.set(instance, Double.parseDouble(cellValue));
                } else if (Boolean.class.equals(fieldType) || boolean.class.equals(fieldType)) {
                    field.set(instance, "1".equals(cellValue) || "true".equalsIgnoreCase(cellValue));
                }
            } catch (NumberFormatException e) {
                log.warn("字段[{}]类型转换失败 (值: '{}')", field.getName(), cellValue);
            }
        } catch (Exception e) {
            log.warn("设置字段[{}]值失败: {}", field.getName(), e.getMessage());
        }
    }
    
    /**
     * 获取单元格的值并转换为字符串
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return null;
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double numericValue = cell.getNumericCellValue();
                    if (numericValue == Math.floor(numericValue)) {
                        return String.valueOf((long) numericValue);
                    } else {
                        return String.valueOf(numericValue);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    return String.valueOf(cell.getNumericCellValue());
                }
            default:
                return null;
        }
    }
}