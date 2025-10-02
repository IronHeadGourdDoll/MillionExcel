package com.example.excel.excel.impl;

import com.example.excel.excel.ExcelHandler;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.multipart.MultipartFile;

import jakarta.servlet.http.HttpServletResponse;
import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * CSV Excel处理器 - 简化版
 * 仅保留基本功能，用于特殊场景
 */
@Slf4j
public class CsvExcelHandler<T> implements ExcelHandler<T> {

    @Override
    public void export(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) {
        try {
            // 确保文件名以.csv结尾
            if (!filename.toLowerCase().endsWith(".csv")) {
                filename = filename + ".csv";
            }
            
            // 设置响应头
            response.setContentType("text/csv; charset=UTF-8");
            response.setCharacterEncoding("UTF-8");
            String encodedFileName = URLEncoder.encode(filename, StandardCharsets.UTF_8.toString());
            response.setHeader("Content-Disposition", "attachment; filename=" + encodedFileName);
            
            // 获取类的所有字段
            Field[] fields = clazz.getDeclaredFields();
            
            try (OutputStreamWriter writer = new OutputStreamWriter(response.getOutputStream(), StandardCharsets.UTF_8)) {
                // 写入UTF-8 BOM标记，确保Excel正确识别编码
                writer.write("\uFEFF");
                
                // 写入表头
                StringBuilder headerLine = new StringBuilder();
                for (int i = 0; i < fields.length; i++) {
                    if (i > 0) {
                        headerLine.append(",");
                    }
                    headerLine.append(escapeCSV(fields[i].getName()));
                }
                writer.write(headerLine.toString());
                writer.write("\n");
                
                // 写入数据
                for (T data : dataList) {
                    StringBuilder dataLine = new StringBuilder();
                    for (int i = 0; i < fields.length; i++) {
                        if (i > 0) {
                            dataLine.append(",");
                        }
                        
                        fields[i].setAccessible(true);
                        Object value = fields[i].get(data);
                        if (value != null) {
                            dataLine.append(escapeCSV(value.toString()));
                        }
                    }
                    writer.write(dataLine.toString());
                    writer.write("\n");
                }
                
                writer.flush();
            }
        } catch (Exception e) {
            log.error("CSV导出失败", e);
            throw new RuntimeException("导出CSV失败", e);
        }
    }

    @Override
    public List<T> importExcel(MultipartFile file, Class<T> clazz) {
        List<T> dataList = new ArrayList<>();
        
        try {
            try (BufferedReader reader = new BufferedReader(new InputStreamReader(file.getInputStream(), StandardCharsets.UTF_8))) {
                // 跳过BOM标记
                reader.mark(1);
                if (reader.read() != '\uFEFF') {
                    reader.reset();
                }
                
                // 读取表头
                String headerLine = reader.readLine();
                if (headerLine == null) {
                    return dataList;
                }
                
                // 解析表头列
                String[] headers = parseCSVLine(headerLine);
                
                // 获取所有字段
                Field[] fields = clazz.getDeclaredFields();
                
                // 创建字段索引映射
                Map<String, Integer> fieldIndexMap = new HashMap<>();
                for (int i = 0; i < headers.length; i++) {
                    fieldIndexMap.put(headers[i].trim(), i);
                }
                
                // 读取数据行
                String line;
                while ((line = reader.readLine()) != null) {
                    // 跳过空行和注释行
                    if (line.trim().isEmpty() || line.trim().startsWith("#")) {
                        continue;
                    }
                    
                    // 解析CSV行
                    String[] values = parseCSVLine(line);
                    
                    // 创建对象实例
                    T instance = clazz.getDeclaredConstructor().newInstance();
                    
                    // 设置字段值 - 根据字段名匹配CSV列
                    for (Field field : fields) {
                        String fieldName = field.getName();
                        if (fieldIndexMap.containsKey(fieldName)) {
                            int colIndex = fieldIndexMap.get(fieldName);
                            if (colIndex < values.length && values[colIndex] != null && !values[colIndex].isEmpty()) {
                                field.setAccessible(true);
                                setFieldValue(field, instance, values[colIndex]);
                            }
                        }
                    }
                    
                    dataList.add(instance);
                }
            }
        } catch (Exception e) {
            log.error("CSV导入失败", e);
            throw new RuntimeException("导入CSV失败", e);
        }
        
        return dataList;
    }
    
    /**
     * 解析CSV行，处理引号和逗号
     */
    private String[] parseCSVLine(String line) {
        List<String> result = new ArrayList<>();
        StringBuilder current = new StringBuilder();
        boolean inQuotes = false;
        
        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);
            
            if (c == '\"') {
                // 处理引号
                if (inQuotes && i + 1 < line.length() && line.charAt(i + 1) == '\"') {
                    // 转义的引号
                    current.append('\"');
                    i++;
                } else {
                    // 开始或结束引号
                    inQuotes = !inQuotes;
                }
            } else if (c == ',' && !inQuotes) {
                // 逗号分隔符（不在引号内）
                result.add(current.toString());
                current = new StringBuilder();
            } else {
                // 普通字符
                current.append(c);
            }
        }
        
        // 添加最后一个字段
        result.add(current.toString());
        
        return result.toArray(new String[0]);
    }
    
    /**
     * 转义CSV字段值
     */
    private String escapeCSV(String value) {
        if (value == null) {
            return "";
        }
        
        // 如果包含逗号、引号或换行符，需要用引号包围
        if (value.contains(",") || value.contains("\"") || value.contains("\n")) {
            // 将引号替换为两个引号（CSV转义规则）
            value = value.replace("\"", "\"\"");
            // 用引号包围
            return "\"" + value + "\"";
        }
        
        return value;
    }
    
    /**
     * 根据字段类型设置值
     */
    private void setFieldValue(Field field, T instance, String value) {
        try {
            Class<?> fieldType = field.getType();
            
            if (String.class.equals(fieldType)) {
                field.set(instance, value);
            } else if (Integer.class.equals(fieldType) || int.class.equals(fieldType)) {
                field.set(instance, Integer.parseInt(value));
            } else if (Long.class.equals(fieldType) || long.class.equals(fieldType)) {
                field.set(instance, Long.parseLong(value));
            } else if (Double.class.equals(fieldType) || double.class.equals(fieldType)) {
                field.set(instance, Double.parseDouble(value));
            } else if (Boolean.class.equals(fieldType) || boolean.class.equals(fieldType)) {
                field.set(instance, Boolean.parseBoolean(value));
            }
            // 其他类型可以根据需要添加
        } catch (Exception e) {
            log.warn("设置字段值失败: {}", field.getName(), e);
        }
    }
}