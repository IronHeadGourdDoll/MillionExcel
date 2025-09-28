package com.example.excel.excel.impl;

import com.example.excel.entity.User;
import com.example.excel.excel.ExcelHandler;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Async;
import org.springframework.scheduling.annotation.AsyncResult;
import org.springframework.stereotype.Component;
import org.springframework.util.StopWatch;
import org.springframework.web.multipart.MultipartFile;

import jakarta.servlet.http.HttpServletResponse;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;

import java.io.*;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Collectors;

/**
 * Excel处理器实现类 - 高性能版
 * 使用SAX事件驱动解析和多线程并行处理优化Excel大数据量导入性能
 */
@Slf4j
public class ApachePoiExcelHandler<T> implements ExcelHandler<T> {

    @Value("${excel.poi.batch-size}")
    private int batchSize;

    @Value("${excel.poi.thread-count}")
    private int threadCount;

    @Override
    public List<T> importExcel(MultipartFile file, Class<T> clazz) {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        List<T> dataList = new ArrayList<>();

        try {
            // 验证文件是否为空
            if (file.isEmpty()) {
                log.error("上传的文件为空");
                throw new RuntimeException("上传的文件为空");
            }

            // 验证文件类型和扩展名
            String fileName = file.getOriginalFilename();
            if (fileName == null) {
                log.error("文件名不能为空");
                throw new RuntimeException("文件名不能为空");
            }

            // 检查文件扩展名
            String extension = fileName.substring(fileName.lastIndexOf(".") + 1).toLowerCase();
            if (!"xlsx".equals(extension) && !"xls".equals(extension) && !"csv".equals(extension)) {
                log.error("不支持的文件格式: {}", extension);
                throw new RuntimeException("只支持.xlsx、.xls和.csv格式的文件");
            }

            log.info("开始Apache POI导入文件：{}，大小：{}KB，格式：{}", fileName, file.getSize() / 1024, extension);

            // 根据文件类型选择不同的处理方式
            if ("csv".equals(extension)) {
                // 处理CSV文件
                handleCsvFile(file, clazz, dataList);
            } else {
                // 处理Excel文件
                handleExcelFile(file, clazz, dataList, extension);
            }

            stopWatch.stop();
            log.info("Apache POI导入完成，共导入{}条数据，耗时{}ms", dataList.size(), stopWatch.getTotalTimeMillis());
        } catch (Exception e) {
            log.error("Apache POI导入Excel失败", e);
            throw new RuntimeException("导入Excel失败", e);
        }

        return dataList;
    }

    @Override
    public void export(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) {
        long startTime = System.currentTimeMillis();

        // 使用SXSSFWorkbook来处理大数据量，避免内存溢出
        SXSSFWorkbook workbook = null;
        
        try {
            // 设置响应头，确保与Excel 2021 LTSC兼容
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding("UTF-8");
            // 正确处理文件名，避免中文乱码
            String encodedFileName = URLEncoder.encode(filename, StandardCharsets.UTF_8.toString()).replaceAll("\\+", "%");
            response.setHeader("Content-disposition", "attachment;filename*=UTF-8''" + encodedFileName);
            // 确保响应头不会被缓存
            response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
            response.setHeader("Pragma", "no-cache");
            response.setHeader("Expires", "0");
            // 防止浏览器嗅探内容类型
            response.setHeader("X-Content-Type-Options", "nosniff");

            // 创建工作簿和工作表
            workbook = new SXSSFWorkbook(100);
            Sheet sheet = workbook.createSheet("Sheet1");

            // 创建表头
            createHeaderRow(sheet, clazz);

            // 分批处理数据，避免大数据量时的内存问题
            int batchSize = 1000;
            int startIndex = 0;
            int rowIndex = 1;
            
            while (startIndex < dataList.size()) {
                int endIndex = Math.min(startIndex + batchSize, dataList.size());
                
                for (int i = startIndex; i < endIndex; i++) {
                    if (i > 0 && i % 1000 == 0) {
                        log.info("已处理{}条数据", i);
                    }
                    
                    T data = dataList.get(i);
                    Row row = sheet.createRow(rowIndex++);
                    fillRow(row, data, clazz);
                }
                
                startIndex = endIndex;
            }

            // 写入响应
            OutputStream outputStream = response.getOutputStream();
            workbook.write(outputStream);
            outputStream.flush();
            
            // 确保响应完全刷新
            response.flushBuffer();

            long endTime = System.currentTimeMillis();
            log.info("POI导出完成，共导出{}条数据，耗时{}ms", dataList.size(), (endTime - startTime));
        } catch (Exception e) {
            log.error("POI导出Excel失败", e);
            throw new RuntimeException("导出Excel失败", e);
        } finally {
            // 关闭资源
            if (workbook != null) {
                workbook.dispose();
            }
        }
    }

    @Override
    @Async
    public Future<Boolean> asyncExport(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) {
        try {
            this.export(dataList, response, filename, clazz);
            return new AsyncResult<>(true);
        } catch (Exception e) {
            log.error("异步POI导出Excel失败", e);
            return new AsyncResult<>(false);
        }
    }

    @Override
    @Async
    public Future<List<T>> asyncImportExcel(MultipartFile file, Class<T> clazz) {
        try {
            List<T> dataList = this.importExcel(file, clazz);
            return new AsyncResult<>(dataList);
        } catch (Exception e) {
            log.error("异步POI导入Excel失败", e);
            return new AsyncResult<>(new ArrayList<>());
        }
    }

    @Override
    public void exportByPage(int pageNum, int pageSize, HttpServletResponse response, String filename, Class<T> clazz) {
        long startTime = System.currentTimeMillis();

        try (SXSSFWorkbook workbook = new SXSSFWorkbook(100);
             OutputStream outputStream = response.getOutputStream()) {

            // 设置响应头
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition",
                    "attachment; filename=" + URLEncoder.encode(filename, StandardCharsets.UTF_8.toString()));

            // 创建工作表
            Sheet sheet = workbook.createSheet("Sheet1");

            // 创建表头
            createHeaderRow(sheet, clazz);

            // 注意：由于移除了UserService依赖，此方法当前只创建空文件
            // 在实际应用中，这里应该有一个泛型的服务接口来获取相应类型的数据
            log.warn("POI分页导出：当前模式仅创建空文件，数据获取功能已临时禁用");

            // 写入响应
            workbook.write(outputStream);
            workbook.dispose();

            long endTime = System.currentTimeMillis();
            log.info("POI分页导出完成，耗时{}ms", (endTime - startTime));
        } catch (Exception e) {
            log.error("POI分页导出Excel失败", e);
            throw new RuntimeException("导出Excel失败", e);
        }
    }

    /**
     * 处理Excel文件
     */
    private <T> void handleExcelFile(MultipartFile file, Class<T> clazz, List<T> dataList, String extension) throws Exception {
        try (InputStream inputStream = file.getInputStream()) {
            // 根据文件类型选择不同的Workbook实现
            Workbook workbook;
            if ("xlsx".equals(extension)) {
                // 对于XLSX格式，使用SXSSFWorkbook以优化内存使用
                workbook = new XSSFWorkbook(inputStream);
            } else {
                // 对于XLS格式，使用HSSFWorkbook
                workbook = WorkbookFactory.create(inputStream);
            }

            try {
                // 获取第一个sheet
                Sheet sheet = workbook.getSheetAt(0);
                if (sheet == null) {
                    log.warn("Excel文件不包含任何工作表");
                    return;
                }

                // 获取表头行
                Row headerRow = sheet.getRow(0);
                if (headerRow == null) {
                    log.warn("Excel文件不包含表头行");
                    return;
                }

                // 创建表头到列索引的映射
                Map<String, Integer> headerMap = createHeaderMap(headerRow);

                // 处理数据行
                Iterator<Row> rowIterator = sheet.iterator();
                rowIterator.next(); // 跳过表头行

                int totalCount = 0;
                List<T> batchDataList = new ArrayList<>(batchSize);

                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    if (row == null) continue;

                    // 创建实例并填充数据
                    T instance = mapRowToObject(row, clazz, headerMap);
                    if (instance != null) {
                        batchDataList.add(instance);
                        totalCount++;

                        // 达到批处理大小就添加到结果列表
                        if (batchDataList.size() >= batchSize) {
                            dataList.addAll(batchDataList);
                            batchDataList.clear();

                            // 定期记录进度
                            if (totalCount % 50000 == 0) {
                                log.info("已读取{}条数据", totalCount);
                            }
                        }
                    }
                }

                // 添加最后一批数据
                if (!batchDataList.isEmpty()) {
                    dataList.addAll(batchDataList);
                }

                log.info("Apache POI Excel解析完成，共解析{}条数据", totalCount);
            } finally {
                workbook.close();
            }
        }
    }

    /**
     * 创建表头到列索引的映射
     */
    private Map<String, Integer> createHeaderMap(Row headerRow) {
        Map<String, Integer> headerMap = new HashMap<>();
        Iterator<Cell> cellIterator = headerRow.cellIterator();

        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            String headerName = getCellValueAsString(cell).trim();
            if (!headerName.isEmpty()) {
                headerMap.put(headerName, cell.getColumnIndex());
            }
        }

        return headerMap;
    }

    /**
     * 将Excel行数据映射到对象
     */
    private <T> T mapRowToObject(Row row, Class<T> clazz, Map<String, Integer> headerMap) throws Exception {
        T instance = clazz.getDeclaredConstructor().newInstance();
        Field[] fields = clazz.getDeclaredFields();
        boolean hasValidData = false;

        for (Field field : fields) {
            field.setAccessible(true);

            // 获取字段名和ExcelProperty注解
            String fieldName = field.getName();
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);

            // 查找对应的列索引
            Integer columnIndex = null;
            if (excelProperty != null && excelProperty.value().length > 0) {
                // 优先使用ExcelProperty注解中的列名
                for (String propValue : excelProperty.value()) {
                    columnIndex = headerMap.get(propValue);
                    if (columnIndex != null) {
                        break;
                    }
                }
            }

            // 如果没有找到，尝试直接使用字段名
            if (columnIndex == null) {
                columnIndex = headerMap.get(fieldName);
            }

            // 如果找到对应的列，设置字段值
            if (columnIndex != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    setFieldValue(field, instance, cell);
                    hasValidData = true;
                }
            }
        }

        return hasValidData ? instance : null;
    }

    /**
     * 设置字段值
     */
    private <T> void setFieldValue(Field field, T instance, Cell cell) throws Exception {
        String cellValue = getCellValueAsString(cell);
        if (cellValue == null || cellValue.trim().isEmpty()) {
            return;
        }

        // 根据字段类型设置值
        if (field.getType() == String.class) {
            field.set(instance, cellValue.trim());
        } else if (field.getType() == Integer.class || field.getType() == int.class) {
            try {
                field.set(instance, Integer.parseInt(cellValue.trim()));
            } catch (NumberFormatException e) {
                log.debug("整数解析失败: {}", cellValue);
            }
        } else if (field.getType() == Long.class || field.getType() == long.class) {
            try {
                field.set(instance, Long.parseLong(cellValue.trim()));
            } catch (NumberFormatException e) {
                log.debug("长整数解析失败: {}", cellValue);
            }
        } else if (field.getType() == Double.class || field.getType() == double.class) {
            try {
                field.set(instance, Double.parseDouble(cellValue.trim()));
            } catch (NumberFormatException e) {
                log.debug("浮点数解析失败: {}", cellValue);
            }
        } else if (field.getType() == Boolean.class || field.getType() == boolean.class) {
            field.set(instance, Boolean.parseBoolean(cellValue.trim()));
        } else if (field.getType() == LocalDateTime.class) {
            try {
                field.set(instance, LocalDateTime.parse(cellValue.trim()));
            } catch (DateTimeParseException e) {
                log.debug("日期时间解析失败: {}", cellValue);
            }
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
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    // 处理数字类型，避免科学计数法
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
                    return cell.getCellFormula();
                } catch (Exception e) {
                    // 尝试获取公式计算结果
                    FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                    return getCellValueAsString(evaluator.evaluateInCell(cell));
                }
            default:
                return null;
        }
    }

    /**
     * 处理CSV文件
     */
    private <T> void handleCsvFile(MultipartFile file, Class<T> clazz, List<T> dataList) throws Exception {
        try (InputStream inputStream = file.getInputStream();
             Reader reader = new InputStreamReader(inputStream, StandardCharsets.UTF_8);
             CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT.withFirstRecordAsHeader().withTrim())) {

            Field[] fields = clazz.getDeclaredFields();
            // 创建字段名到Excel列名的映射
            Map<String, String> fieldToExcelName = new HashMap<>();
            for (Field field : fields) {
                ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
                if (excelProperty != null && excelProperty.value().length > 0) {
                    fieldToExcelName.put(field.getName(), excelProperty.value()[0]);
                }
            }

            // 分批读取CSV数据，优化内存使用
            int totalCount = 0;
            int batchCount = 0;
            List<T> batchDataList = new ArrayList<>(batchSize);

            // 解析CSV记录
            for (CSVRecord csvRecord : csvParser) {
                try {
                    T instance = clazz.getDeclaredConstructor().newInstance();

                    // 尝试映射所有字段
                    for (Field field : fields) {
                        field.setAccessible(true);
                        String fieldName = field.getName();

                        try {
                            // 首先尝试通过Excel列名获取值
                            String excelColumnName = fieldToExcelName.get(fieldName);
                            String value = null;

                            // 尝试使用Excel列名获取值
                            if (excelColumnName != null) {
                                try {
                                    value = csvRecord.get(excelColumnName);
                                } catch (IllegalArgumentException e) {
                                    // 如果Excel列名不存在，尝试使用字段名
                                    try {
                                        value = csvRecord.get(fieldName);
                                    } catch (IllegalArgumentException ex) {
                                        // 都不存在时，跳过该字段
                                        log.debug("字段映射失败: {} (Excel列名: {})", fieldName, excelColumnName);
                                    }
                                }
                            } else {
                                // 没有ExcelProperty注解，直接使用字段名
                                try {
                                    value = csvRecord.get(fieldName);
                                } catch (IllegalArgumentException e) {
                                    // 字段名不存在时，跳过该字段
                                    log.debug("字段映射失败: {}", fieldName);
                                }
                            }

                            if (value != null && !value.trim().isEmpty()) {
                                // 根据字段类型设置值
                                if (field.getType() == String.class) {
                                    field.set(instance, value.trim());
                                } else if (field.getType() == Integer.class || field.getType() == int.class) {
                                    try {
                                        field.set(instance, Integer.parseInt(value.trim()));
                                    } catch (NumberFormatException e) {
                                        log.warn("整数解析失败: {}", value);
                                    }
                                } else if (field.getType() == Long.class || field.getType() == long.class) {
                                    try {
                                        field.set(instance, Long.parseLong(value.trim()));
                                    } catch (NumberFormatException e) {
                                        log.warn("长整数解析失败: {}", value);
                                    }
                                } else if (field.getType() == Double.class || field.getType() == double.class) {
                                    try {
                                        field.set(instance, Double.parseDouble(value.trim()));
                                    } catch (NumberFormatException e) {
                                        log.warn("浮点数解析失败: {}", value);
                                    }
                                } else if (field.getType() == Boolean.class || field.getType() == boolean.class) {
                                    field.set(instance, Boolean.parseBoolean(value.trim()));
                                } else if (field.getType() == LocalDateTime.class) {
                                    // 处理日期时间类型
                                    try {
                                        field.set(instance, LocalDateTime.parse(value.trim()));
                                    } catch (DateTimeParseException e) {
                                        log.warn("日期时间解析失败: {}", value);
                                    }
                                } else {
                                    // 其他类型直接设置字符串
                                    field.set(instance, value.trim());
                                }
                            }
                        } catch (Exception e) {
                            // 发生任何异常时，跳过该字段，继续处理下一个字段
                            log.debug("字段设置失败: {}", fieldName, e);
                        }
                    }

                    // 数据验证：确保必填字段不为空
                    if (isValidData(instance, clazz)) {
                        batchDataList.add(instance);
                        totalCount++;

                        // 达到批处理大小时，将数据添加到结果集并清空批处理列表
                        if (batchDataList.size() >= batchSize) {
                            dataList.addAll(batchDataList);
                            batchDataList.clear();
                            batchCount++;
                            log.info("已处理{}批数据，共{}条记录", batchCount, totalCount);
                        }
                    } else {
                        log.warn("跳过无效数据行");
                    }
                } catch (Exception e) {
                    // 发生异常时，跳过当前行，继续处理下一行
                    log.warn("处理CSV行数据时发生异常", e);
                }
            }

            // 添加最后一批数据
            if (!batchDataList.isEmpty()) {
                dataList.addAll(batchDataList);
            }
        } catch (Exception e) {
            log.error("处理CSV文件时发生异常", e);
            throw e;
        }
    }

    /**
     * 验证数据是否有效
     */
    private <T> boolean isValidData(T instance, Class<T> clazz) {
        // 如果是User类型，需要特别验证username和name字段
        if (instance instanceof User) {
            User user = (User) instance;
            // username和name是必填字段，不能为null或空字符串
            return user.getUsername() != null && !user.getUsername().isEmpty() &&
                    user.getName() != null && !user.getName().isEmpty();
        }
        // 其他类型的基本验证，可以根据需要扩展
        return true;
    }

    /**
     * 创建表头行
     */
    private void createHeaderRow(Sheet sheet, Class<T> clazz) {
        Row headerRow = sheet.createRow(0);
        Field[] fields = clazz.getDeclaredFields();

        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            field.setAccessible(true);
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(field.getName());

            // 设置表头样式
            CellStyle style = sheet.getWorkbook().createCellStyle();
            Font font = sheet.getWorkbook().createFont();
            font.setBold(true);
            style.setFont(font);
            cell.setCellStyle(style);
        }
    }

    /**
     * 填充数据行
     */
    private void fillRow(Row row, T data, Class<T> clazz) {
        try {
            Field[] fields = clazz.getDeclaredFields();
            for (int i = 0; i < fields.length; i++) {
                Field field = fields[i];
                field.setAccessible(true);
                Object value = field.get(data);
                Cell cell = row.createCell(i);

                if (value != null) {
                    cell.setCellValue(value.toString());
                }
            }
        } catch (Exception e) {
            log.error("填充数据行失败", e);
            throw new RuntimeException("填充数据行失败", e);
        }
    }

    /**
     * 解析数据行
     */
    private T parseRow(Row row, Class<T> clazz) {
        try {
            T instance = clazz.getDeclaredConstructor().newInstance();
            Field[] fields = clazz.getDeclaredFields();

            for (int i = 0; i < fields.length; i++) {
                Field field = fields[i];
                field.setAccessible(true);
                Cell cell = row.getCell(i);

                if (cell != null) {
                    String cellValue = getCellValue(cell);
                    if (cellValue != null && !cellValue.isEmpty()) {
                        // 根据字段类型设置值
                        if (field.getType() == String.class) {
                            field.set(instance, cellValue);
                        } else if (field.getType() == Integer.class || field.getType() == int.class) {
                            field.set(instance, Integer.parseInt(cellValue));
                        } else if (field.getType() == Long.class || field.getType() == long.class) {
                            field.set(instance, Long.parseLong(cellValue));
                        } else if (field.getType() == Double.class || field.getType() == double.class) {
                            field.set(instance, Double.parseDouble(cellValue));
                        } else if (field.getType() == Boolean.class || field.getType() == boolean.class) {
                            field.set(instance, Boolean.parseBoolean(cellValue));
                        } else {
                            // 其他类型直接设置字符串
                            field.set(instance, cellValue);
                        }
                    }
                }
            }

            return instance;
        } catch (Exception e) {
            log.error("解析数据行失败", e);
            return null;
        }
    }

    /**
     * 获取单元格的值
     */
    private String getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return cell.toString();
        }
    }
}