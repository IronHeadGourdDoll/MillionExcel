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
import java.lang.management.ManagementFactory;
import java.lang.management.MemoryMXBean;
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

    // 线程池配置
    private static final int CORE_POOL_SIZE = Runtime.getRuntime().availableProcessors() * 2;
    private static final ExecutorService executorService = Executors.newFixedThreadPool(CORE_POOL_SIZE);
    
    @Override
    public void export(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        MemoryMXBean memoryMXBean = ManagementFactory.getMemoryMXBean();
        
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
            
            // 内存检查
            long usedMemory = memoryMXBean.getHeapMemoryUsage().getUsed() / (1024 * 1024);
            long maxMemory = memoryMXBean.getHeapMemoryUsage().getMax() / (1024 * 1024);
            if (usedMemory > maxMemory * 0.8) {
                log.warn("内存使用率超过80%，当前使用: {}MB/{}MB", usedMemory, maxMemory);
            }
            
            try {
                ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream(), clazz)
                        .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                        .autoTrim(true)
                        .inMemory(false)
                        .build();
                
                WriteSheet writeSheet = EasyExcel.writerSheet("Sheet1")
                        .build();
                
                // 动态调整批处理大小并启用多线程写入
                int batchSize = calculateBatchSize(dataList.size());
                int totalBatches = (dataList.size() + batchSize - 1) / batchSize;
                List<Future<?>> futures = new ArrayList<>(totalBatches);
                
                for (int i = 0; i < dataList.size(); i += batchSize) {
                    final int start = i;
                    final int end = Math.min(i + batchSize, dataList.size());
                    List<T> batch = dataList.subList(start, end);
                    
                    futures.add(executorService.submit(() -> {
                        excelWriter.write(batch, writeSheet);
                        return null;
                    }));
                    
                    // 控制并发任务数量
                    if (futures.size() >= CORE_POOL_SIZE * 2) {
                        waitForFutures(futures);
                    }
                    excelWriter.write(batch, writeSheet);
                    
                    // 每10批或最后一批记录日志
                    if ((i / batchSize) % 10 == 0 || end == dataList.size()) {
                        log.info("导出进度: {}/{} ({}%)", 
                            (i / batchSize) + 1, 
                            totalBatches,
                            Math.round(((i + batchSize) * 100.0) / dataList.size()));
                    }
                    
                    // 批处理间执行GC
                    if (i > 0 && i % (batchSize * 10) == 0) {
                        System.gc();
                    }
                }
                
                // 等待所有任务完成
                waitForFutures(futures);
                
                // 刷新并关闭资源
                excelWriter.finish();
                response.flushBuffer();
            } finally {
                // 确保资源正确关闭
            }

            // 强制清理临时文件
            System.gc();
            
            stopWatch.stop();
            log.info("Excel导出完成，共导出{}条数据，总耗时{}ms，平均速度:{}/s", 
                dataList.size(), 
                stopWatch.getTotalTimeMillis(),
                (int)(dataList.size() * 1000.0 / stopWatch.getTotalTimeMillis()));
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
            if (file == null || file.isEmpty()) {
                throw new IllegalArgumentException("上传文件不能为空");
            }
            
            log.info("开始导入文件：{}，大小：{}KB", file.getOriginalFilename(), file.getSize() / 1024);
            
            // 使用EasyExcel读取数据
            EasyExcel.read(file.getInputStream(), clazz, new DataReadListener<>(dataList))
                    .sheet()
                    .doRead();
            
            stopWatch.stop();
            log.info("Excel导入完成，共导入{}条数据，耗时{}ms", dataList.size(), stopWatch.getTotalTimeMillis());
        } catch (IllegalArgumentException e) {
            log.error("Excel导入参数错误", e);
            throw e;
        } catch (Exception e) {
            log.error("Excel导入失败", e);
            throw new RuntimeException("导入Excel失败", e);
        }
        
        return dataList;
    }
    
    /**
     * 根据数据量动态计算批处理大小
     */
    private int calculateBatchSize(int totalSize) {
        if (totalSize <= 10_000) {
            return 1_000;
        } else if (totalSize <= 100_000) {
            return 5_000;
        } else if (totalSize <= 500_000) {
            return 20_000;
        } else if (totalSize <= 1_000_000) {
            return 50_000;
        } else {
            return 100_000;
        }
    }
    
    /**
     * 等待所有Future任务完成
     */
    private void waitForFutures(List<Future<?>> futures) {
        for (Future<?> future : futures) {
            try {
                future.get();
            } catch (Exception e) {
                log.error("导出任务执行失败", e);
            }
        }
        futures.clear();
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