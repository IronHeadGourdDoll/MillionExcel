package com.example.excel.excel.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.example.excel.excel.ExcelHandler;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.formula.functions.T;
import org.springframework.stereotype.Component;
import org.springframework.util.StopWatch;
import org.springframework.web.multipart.MultipartFile;

import jakarta.servlet.http.HttpServletResponse;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

/**
 * 统一的Excel处理器实现，使用EasyExcel
 * 简化版：移除了多余的配置和复杂逻辑，保留核心功能
 */
@Slf4j
@Component
public class EasyExcelHandler<T> implements ExcelHandler<T> {

    @Override
    public void export(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) {
        long startTime = System.currentTimeMillis();
        
        try {
            // 确保文件名以.xlsx结尾
            if (!filename.toLowerCase().endsWith(".xlsx")) {
                filename = filename + ".xlsx";
            }
            
            // 设置响应头
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding("UTF-8");
            String encodedFileName = URLEncoder.encode(filename, StandardCharsets.UTF_8.toString()).replaceAll("\\+", "%");
            response.setHeader("Content-disposition", "attachment;filename*=UTF-8''" + encodedFileName);
            response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
            response.setHeader("Pragma", "no-cache");
            response.setHeader("Expires", "0");
            
            try {
                ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream(), clazz)
                        .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                        .autoTrim(true)
                        .inMemory(false)
                        .build();
                
                WriteSheet writeSheet = EasyExcel.writerSheet("Sheet1")
                        .build();
                
                // 单线程顺序写入，避免多线程安全问题
                int batchSize = 200000; // 每批处理20万条
                for (int i = 0; i < dataList.size(); i += batchSize) {
                    final int start = i;
                    final int end = Math.min(i + batchSize, dataList.size());
                    List<T> batch = dataList.subList(start, end);
                    excelWriter.write(batch, writeSheet);
                    log.debug("已导出 {} 到 {} 条数据", start, end);
                }
                
                // 刷新并关闭资源
                excelWriter.finish();
                response.flushBuffer();
            } finally {
                // 确保资源正确关闭
            }

            // 强制清理临时文件
            System.gc();
            
            long endTime = System.currentTimeMillis();
            log.info("Excel导出完成，共导出{}条数据，耗时{}ms", dataList.size(), (endTime - startTime));
        } catch (Exception e) {
            log.error("Excel导出失败", e);
            throw new RuntimeException("导出Excel失败", e);
        }
    }

    @Override
    public List<T> importExcel(MultipartFile file, Class<T> clazz) {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        List<T> dataList = new ArrayList<>();
        
        try {
            log.info("开始导入文件：{}，大小：{}KB", file.getOriginalFilename(), file.getSize() / 1024);
            
            // 使用EasyExcel读取数据
            EasyExcel.read(file.getInputStream(), clazz, new DataReadListener<>(dataList))
                    .sheet()
                    .doRead();
            
            stopWatch.stop();
            log.info("Excel导入完成，共导入{}条数据，耗时{}ms", dataList.size(), stopWatch.getTotalTimeMillis());
        } catch (Exception e) {
            log.error("Excel导入失败", e);
            throw new RuntimeException("导入Excel失败", e);
        }
        
        return dataList;
    }

    /**
     * 数据读取监听器，用于EasyExcel读取数据
     * 简化版：保留批处理功能，提高大数据量处理性能
     */
    private static class DataReadListener<T> extends AnalysisEventListener<T> {
        
        private final List<T> dataList;
        private static final int BATCH_COUNT = 5000; // 批处理大小
        private final List<T> batchList = new ArrayList<>(BATCH_COUNT);
        
        public DataReadListener(List<T> dataList) {
            this.dataList = dataList;
        }
        
        @Override
        public void invoke(T data, AnalysisContext context) {
            batchList.add(data);
            
            // 达到批处理数量就保存一次，减少内存占用
            if (batchList.size() >= BATCH_COUNT) {
                saveData();
            }
        }
        
        @Override
        public void doAfterAllAnalysed(AnalysisContext context) {
            // 保存最后一批数据
            saveData();
            log.info("Excel读取完成，共读取{}条数据", dataList.size());
        }
        
        private void saveData() {
            if (!batchList.isEmpty()) {
                dataList.addAll(batchList);
                batchList.clear();
            }
        }
        
        @Override
        public void onException(Exception exception, AnalysisContext context) {
            log.error("读取Excel数据时发生异常：第{}行", context.readRowHolder().getRowIndex(), exception);
        }
    }
}