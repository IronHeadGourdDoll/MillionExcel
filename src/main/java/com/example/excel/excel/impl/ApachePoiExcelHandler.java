package com.example.excel.excel.impl;

import com.example.excel.excel.ExcelHandler;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

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
        List<T> dataList = new ArrayList<>();
        
        try {
            // 创建临时文件
            java.nio.file.Path tempFile = java.nio.file.Files.createTempFile("excel-import", ".xlsx");
            try {
                // 将上传文件保存到临时文件
                file.transferTo(tempFile);
                
                // 使用安全模式从临时文件加载
                try (org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(tempFile.toFile(), null, true)) {
                    // 获取第一个sheet
                    Sheet sheet = workbook.getSheetAt(0);
                    
                    // 获取所有字段
                    Field[] fields = clazz.getDeclaredFields();
                    
                    // 获取行迭代器
                    Iterator<Row> rowIterator = sheet.iterator();
                    
                    // 如果没有行，直接返回空列表
                    if (!rowIterator.hasNext()) {
                        return dataList;
                    }
                    
                    // 读取表头行，建立表头与列索引的映射
                    Row headerRow = rowIterator.next();
                    Map<String, Integer> headerMap = new HashMap<>();
                    for (Cell cell : headerRow) {
                        String headerName = getCellValueAsString(cell);
                        if (headerName != null && !headerName.trim().isEmpty()) {
                            headerMap.put(headerName.trim(), cell.getColumnIndex());
                            log.debug("表头映射: [{}] -> 列{}", headerName, cell.getColumnIndex());
                        }
                    }
                    
                    // 遍历数据行
                    while (rowIterator.hasNext()) {
                        Row row = rowIterator.next();
                        T instance = clazz.getDeclaredConstructor().newInstance();
                        
                        // 遍历字段，根据表头映射设置值
                        for (Field field : fields) {
                            // 跳过id字段，由后端生成
                            if ("id".equals(field.getName())) {
                                continue;
                            }
                            
                            // 查找当前字段在Excel中的列索引
                            Integer colIndex = headerMap.get(field.getName());
                            if (colIndex == null) {
                                // 尝试大小写不敏感匹配
                                for (Map.Entry<String, Integer> entry : headerMap.entrySet()) {
                                    if (entry.getKey().equalsIgnoreCase(field.getName())) {
                                        colIndex = entry.getValue();
                                        break;
                                    }
                                }
                            }
                            
                            // 如果找到对应的列，设置字段值
                            if (colIndex != null) {
                                Cell cell = row.getCell(colIndex);
                                if (cell != null) {
                                    field.setAccessible(true);
                                    setFieldValue(field, instance, cell);
                                }
                            }
                        }
                        
                        dataList.add(instance);
                    }
                }
            } catch (IOException e) {
                throw new RuntimeException(e);
            } catch (IllegalStateException e) {
                throw new RuntimeException(e);
            } catch (SecurityException e) {
                throw new RuntimeException(e);
            } catch (InstantiationException e) {
                throw new RuntimeException(e);
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            } catch (IllegalArgumentException e) {
                throw new RuntimeException(e);
            } catch (InvocationTargetException e) {
                throw new RuntimeException(e);
            } catch (NoSuchMethodException e) {
                throw new RuntimeException(e);
            }
        } catch (Exception e) {
            log.error("Apache POI导入Excel失败", e);
            throw new RuntimeException("导入Excel失败", e);
        }
        
        return dataList;
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