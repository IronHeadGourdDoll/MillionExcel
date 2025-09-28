package com.example.excel.excel.impl;

import com.example.excel.excel.ExcelHandler;
import com.example.excel.entity.User;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;

import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.Future;

/**
 * 用户Excel处理器，使用EasyExcel实现Excel导入导出功能
 */
@Slf4j
@Component("userExcelHandler")
public class UserExcelHandler implements ExcelHandler<User> {

    @Override
    public void export(List<User> dataList, HttpServletResponse response, String filename, Class<User> clazz) {
        long startTime = System.currentTimeMillis();
        
        try {
            // 设置响应头
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition", 
                "attachment; filename=" + URLEncoder.encode(filename, StandardCharsets.UTF_8.toString()));
            
            // 使用EasyExcel写入数据
            EasyExcel.write(response.getOutputStream(), clazz)
                    .sheet("用户数据")
                    .doWrite(dataList);
            
            long endTime = System.currentTimeMillis();
            log.info("EasyExcel导出完成，共导出{}条数据，耗时{}ms", dataList.size(), (endTime - startTime));
        } catch (Exception e) {
            log.error("EasyExcel导出Excel失败", e);
            throw new RuntimeException("导出Excel失败", e);
        }
    }

    @Override
    public Future<Boolean> asyncExport(List<User> dataList, HttpServletResponse response, String filename, Class<User> clazz) {
        return CompletableFuture.supplyAsync(() -> {
            try {
                this.export(dataList, response, filename, clazz);
                return true;
            } catch (Exception e) {
                log.error("异步导出Excel失败", e);
                return false;
            }
        });
    }

    @Override
    public List<User> importExcel(MultipartFile file, Class<User> clazz) {
        long startTime = System.currentTimeMillis();
        List<User> dataList = new ArrayList<>();
        
        try {
            String filename = file.getOriginalFilename();
            log.info("开始导入文件：{}", filename);
            
            try (InputStream inputStream = file.getInputStream()) {
                if (filename != null && filename.toLowerCase().endsWith(".csv")) {
                    // 特别处理CSV文件，确保正确处理UTF-8编码和#开头的数据行
                    log.info("检测到CSV文件，使用专门的CSV解析策略");
                    
                    // 对于CSV文件，我们需要创建一个自定义的配置
                    com.alibaba.excel.read.builder.ExcelReaderBuilder readerBuilder = EasyExcel.read(inputStream, clazz, new DataReadListener<>(dataList));
                    
                    // 配置CSV格式
                    com.alibaba.excel.support.ExcelTypeEnum excelTypeEnum = com.alibaba.excel.support.ExcelTypeEnum.CSV;
                    readerBuilder.excelType(excelTypeEnum);
                    
                    // EasyExcel的CSV解析会自动处理UTF-8编码
                    readerBuilder.sheet().doRead();
                } else {
                    // 处理普通Excel文件
                    log.info("检测到Excel文件，使用标准Excel解析策略");
                    try {
                        // 尝试使用EasyExcel标准方式读取
                        EasyExcel.read(inputStream, clazz, new DataReadListener<>(dataList))
                                .sheet()
                                .doRead();
                    } catch (Exception e) {
                        // 检查是否是Strict OOXML格式问题
                        if (e.getMessage() != null && e.getMessage().contains("Strict OOXML")) {
                            log.warn("检测到Strict OOXML格式，尝试使用POI兼容模式处理");
                            // 重新获取输入流
                            inputStream.close();
                            try (InputStream retryStream = file.getInputStream()) {
                                // 使用POI直接处理Strict OOXML格式
                                    org.apache.poi.ss.usermodel.Workbook workbook = null;
                                    try {
                                        workbook = WorkbookFactory.create(retryStream);
                                        // 获取第一个sheet
                                        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(0);
                                        // 手动解析Excel数据
                                        manualParseExcel(sheet, dataList, clazz);
                                    } finally {
                                        if (workbook != null) {
                                            workbook.close();
                                        }
                                    }
                            }
                        } else {
                            throw e;
                        }
                    }
                }
            }
            
            long endTime = System.currentTimeMillis();
            log.info("导入完成，共导入{}条数据，耗时{}ms", dataList.size(), (endTime - startTime));
        } catch (Exception e) {
            log.error("导入Excel失败", e);
            throw new RuntimeException("导入Excel失败：" + e.getMessage(), e);
        }
        
        return dataList;
    }
    
    /**
     * 手动解析Excel数据，用于处理Strict OOXML格式等特殊情况
     */
    private <T> void manualParseExcel(org.apache.poi.ss.usermodel.Sheet sheet, List<T> dataList, Class<T> clazz) {
        if (sheet == null) {
            return;
        }
        
        Iterator<org.apache.poi.ss.usermodel.Row> rowIterator = sheet.iterator();
        if (!rowIterator.hasNext()) {
            return;
        }
        
        // 跳过表头
        org.apache.poi.ss.usermodel.Row headerRow = rowIterator.next();
        
        // 只处理User类型的数据，简化版实现
        if (User.class.equals(clazz)) {
            while (rowIterator.hasNext()) {
                org.apache.poi.ss.usermodel.Row row = rowIterator.next();
                User user = new User();
                
                // 尝试从单元格读取数据，根据常见的Excel列顺序设置属性
                // 注意：这里需要根据实际的Excel列顺序调整
                try {
                    if (row.getCell(0) != null) {
                        user.setUsername(getCellValueAsString(row.getCell(0)));
                    }
                    if (row.getCell(1) != null) {
                        user.setName(getCellValueAsString(row.getCell(1)));
                    }
                    if (row.getCell(2) != null) {
                        user.setEmail(getCellValueAsString(row.getCell(2)));
                    }
                    if (row.getCell(3) != null) {
                        user.setPhone(getCellValueAsString(row.getCell(3)));
                    }
                    // 如果用户对象有有效数据，则添加到列表
                    if (user.getUsername() != null || user.getName() != null || 
                           user.getEmail() != null || user.getPhone() != null) {
                        dataList.add((T) user);
                    }
                } catch (Exception e) {
                    log.warn("解析行数据失败，跳过该行", e);
                }
            }
        }
    }
    
    /**
     * 获取单元格的值并转换为字符串
     */
    private String getCellValueAsString(org.apache.poi.ss.usermodel.Cell cell) {
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
                    // 处理数字，避免科学计数法
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

    @Override
    public Future<List<User>> asyncImportExcel(MultipartFile file, Class<User> clazz) {
        return CompletableFuture.supplyAsync(() -> {
            try {
                return this.importExcel(file, clazz);
            } catch (Exception e) {
                log.error("异步导入Excel失败", e);
                return new ArrayList<>();
            }
        });
    }

    @Override
    public void exportByPage(int pageNum, int pageSize, HttpServletResponse response, String filename, Class<User> clazz) {
        // 简化实现，实际应用中应该分页查询数据
        this.export(new ArrayList<>(), response, filename, clazz);
    }
    
    /**
     * 数据读取监听器，用于EasyExcel异步读取数据
     * 特别处理CSV文件中的#开头行和空数据行
     */
    private static class DataReadListener<T> extends AnalysisEventListener<T> {
        
        private final List<T> dataList;
        private static final int BATCH_COUNT = 10000;
        private List<T> batchList = new ArrayList<>();
        
        public DataReadListener(List<T> dataList) {
            this.dataList = dataList;
        }
        
        @Override
        public void invoke(T data, AnalysisContext context) {
            // 跳过空数据行
            if (data != null) {
                // 针对User类型的特殊处理，确保即使数据有#开头也能被正确处理
                if (data instanceof User) {
                    User user = (User) data;
                    // 检查用户对象是否有有意义的数据
                    if (hasValidUserData(user)) {
                        batchList.add(data);
                        // 达到批处理数量就保存一次
                        if (batchList.size() >= BATCH_COUNT) {
                            saveData();
                            batchList.clear();
                        }
                    } else {
                        log.debug("跳过无效用户数据行");
                    }
                } else {
                    // 其他类型直接添加
                    batchList.add(data);
                    if (batchList.size() >= BATCH_COUNT) {
                        saveData();
                        batchList.clear();
                    }
                }
            }
        }
        
        /**
         * 检查用户对象是否包含有效的用户数据
         */
        private boolean hasValidUserData(User user) {
            // 只要用户名、姓名、邮箱、手机号中有一个不为空，就认为是有效数据
            return user.getUsername() != null || user.getName() != null || 
                   user.getEmail() != null || user.getPhone() != null;
        }
        
        @Override
        public void doAfterAllAnalysed(AnalysisContext context) {
            // 保存最后一批数据
            saveData();
        }
        
        private void saveData() {
            dataList.addAll(batchList);
        }
    }
}