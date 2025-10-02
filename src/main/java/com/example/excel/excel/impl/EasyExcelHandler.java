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

import java.io.UnsupportedEncodingException;
import java.lang.management.ManagementFactory;
import java.lang.management.MemoryMXBean;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.locks.ReentrantLock;
import java.util.function.Function;
import java.util.stream.IntStream;

/**
 * 统一的Excel处理器实现，使用EasyExcel
 * 简化版：移除了多余的配置和复杂逻辑，保留核心功能
 */
@Slf4j
@Component
public class EasyExcelHandler<T> implements ExcelHandler<T> {

    // 线程池配置 - 优化线程池参数以适应大数据量导出
    private static final int CORE_POOL_SIZE = Runtime.getRuntime().availableProcessors() * 2;
    private static final ExecutorService exportExecutor = new ThreadPoolExecutor(
            CORE_POOL_SIZE,
            CORE_POOL_SIZE * 2,
            60L,
            TimeUnit.SECONDS,
            new LinkedBlockingQueue<>(100),
            r -> {
                Thread t = new Thread(r, "excel-export-thread-");
                t.setDaemon(true);
                t.setPriority(Thread.NORM_PRIORITY);
                return t;
            },
            new ThreadPoolExecutor.CallerRunsPolicy() // 当队列满时，在调用者线程执行
    );
    
    // EasyExcel专用配置
    private static final int EXCEL_MAX_ROWS = 1048576; // Excel 2007+ 最大行数
    private static final int DEFAULT_BATCH_SIZE = 50000; // 默认批处理大小
    private static final int MAX_BATCH_SIZE = 200000; // 最大批处理大小
    
    @Override
    public void export(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) throws UnsupportedEncodingException {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();

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

            // 检查数据量是否超过Excel最大行数限制
            if (dataList.size() >= EXCEL_MAX_ROWS) {
                throw new IllegalArgumentException("数据量超出Excel最大行数限制(" + EXCEL_MAX_ROWS + "行)");
            }
            log.info("准备导出数据，总条数：{}", dataList.size());

            // 导出优化策略选择
            if (dataList.size() > 50000) {
                // 大数据量使用多线程预处理+单线程写入模式
                exportLargeDataWithMultiThread(dataList, response, clazz);
            } else {
                // 小数据量使用优化的顺序写入
                exportSmallDataWithOptimizedWrite(dataList, response, clazz);
            }

            stopWatch.stop();
            log.info("Excel导出完成，共导出{}条数据，总耗时{}ms，平均速度:{}/s",
                    dataList.size(),
                    stopWatch.getTotalTimeMillis(),
                    (int) (dataList.size() * 1000.0 / stopWatch.getTotalTimeMillis()));
        } catch (Exception e) {
            log.error("Excel导出失败", e);
            throw new RuntimeException("导出Excel失败", e);
        }
    }
    
    /**
     * 大数据量导出 - 多线程预处理+单线程写入模式
     */
    private void exportLargeDataWithMultiThread(List<T> dataList, HttpServletResponse response, Class<T> clazz) throws Exception {
        // 配置EasyExcel为高性能模式
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream(), clazz)
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                .autoTrim(true)
                .inMemory(false) // 使用磁盘缓存
                .build();
        
        WriteSheet writeSheet = EasyExcel.writerSheet("Sheet1")
                .build();
        
        // 计算最优批处理大小
        int batchSize = calculateOptimalBatchSize(dataList.size());
        int totalBatches = (dataList.size() + batchSize - 1) / batchSize;
        log.info("大数据量导出配置：每批{}条，共{}批，线程数：{}", 
                batchSize, totalBatches, CORE_POOL_SIZE);
        
        // 创建阻塞队列用于存储预处理的数据
        BlockingQueue<List<T>> processedDataQueue = new LinkedBlockingQueue<>(CORE_POOL_SIZE * 2);
        
        // 线程同步控制
        CountDownLatch completionLatch = new CountDownLatch(CORE_POOL_SIZE);
        AtomicInteger processedBatches = new AtomicInteger(0);
        AtomicBoolean exportFailed = new AtomicBoolean(false);
        ReentrantLock writerLock = new ReentrantLock();
        
        // 启动数据预处理线程
        for (int i = 0; i < CORE_POOL_SIZE; i++) {
            final int threadId = i;
            exportExecutor.submit(() -> {
                try {
                    while (true) {
                        int batchIndex = processedBatches.getAndIncrement();
                        if (batchIndex >= totalBatches || exportFailed.get()) {
                            break;
                        }
                        
                        int start = batchIndex * batchSize;
                        int end = Math.min(start + batchSize, dataList.size());
                        
                        // 预处理数据（复制、转换等）
                        List<T> batch = new ArrayList<>(dataList.subList(start, end));
                        
                        // 将预处理后的数据放入队列
                        processedDataQueue.put(batch);
                        
                        // 每处理10批记录一次进度
                        if (batchIndex % 10 == 0) {
                            log.info("数据预处理进度: 线程{} - {}/{}批", threadId, batchIndex, totalBatches);
                        }
                    }
                } catch (Exception e) {
                    log.error("数据预处理线程{}发生异常", threadId, e);
                    exportFailed.set(true);
                } finally {
                    completionLatch.countDown();
                }
            });
        }
        
        // 启动单线程写入Excel
        int writtenBatches = 0;
        while (writtenBatches < totalBatches && !exportFailed.get()) {
            List<T> batchData = processedDataQueue.poll(2, TimeUnit.SECONDS);
            if (batchData != null) {
                writerLock.lock();
                try {
                    excelWriter.write(batchData, writeSheet);
                } catch (Exception e) {
                    log.error("写入第{}批数据失败", writtenBatches, e);
                    exportFailed.set(true);
                } finally {
                    writerLock.unlock();
                }
                
                writtenBatches++;
                
                // 记录写入进度
                if (writtenBatches % 5 == 0 || writtenBatches == totalBatches) {
                    long progress = Math.min(100, Math.round((writtenBatches * 100.0) / totalBatches));
                    log.info("Excel写入进度: {}/{}批 ({}%)", writtenBatches, totalBatches, progress);
                }
                
                // 定期释放内存
                if (writtenBatches % 10 == 0) {
                    batchData.clear();
                    System.gc();
                }
            }
            
            // 检查预处理是否已完成且队列已清空
            if (completionLatch.getCount() == 0 && processedDataQueue.isEmpty()) {
                break;
            }
        }
        
        // 确保资源释放
        excelWriter.finish();
        response.flushBuffer();
        processedDataQueue.clear();
        
        if (exportFailed.get()) {
            throw new RuntimeException("Excel导出过程中发生错误");
        }
    }
    
    /**
     * 小数据量导出 - 优化的顺序写入模式
     */
    private void exportSmallDataWithOptimizedWrite(List<T> dataList, HttpServletResponse response, Class<T> clazz) throws Exception {
        // 配置EasyExcel为高性能模式
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream(), clazz)
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                .autoTrim(true)
                .inMemory(false) // 使用磁盘缓存
                .build();
        
        WriteSheet writeSheet = EasyExcel.writerSheet("Sheet1")
                .build();
        
        // 计算批处理大小
        int batchSize = calculateOptimalBatchSize(dataList.size());
        int totalBatches = (dataList.size() + batchSize - 1) / batchSize;
        log.info("小数据量导出配置：每批{}条，共{}批", batchSize, totalBatches);
        
        // 优化的顺序写入
        for (int batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
            int start = batchIndex * batchSize;
            int end = Math.min(start + batchSize, dataList.size());
            List<T> batch = dataList.subList(start, end);
            
            // 直接写入，减少对象创建
            excelWriter.write(batch, writeSheet);
            
            // 记录进度
            if (batchIndex % 5 == 0 || batchIndex == totalBatches - 1) {
                long progress = Math.min(100, Math.round((end * 100.0) / dataList.size()));
                log.info("导出进度: {}/{} ({}%)", 
                        batchIndex + 1, totalBatches,
                        progress);
            }
        }
        
        // 完成导出
        excelWriter.finish();
        response.flushBuffer();
    }
    
    /**
     * 根据数据量和系统资源动态计算最优批处理大小
     */
    private int calculateOptimalBatchSize(int totalRows) {
        // 获取系统资源信息
        Runtime runtime = Runtime.getRuntime();
        long maxMemory = runtime.maxMemory();
        long freeMemory = runtime.freeMemory();
        long availableMemory = freeMemory + (maxMemory - runtime.totalMemory());
        
        // 预估单条数据大小（字节）
        // 这里根据实际业务情况调整，此为经验值
        int estimatedRowSizeBytes = 1024; // 假设每行约1KB
        
        // 基于可用内存计算批处理大小（使用30%的可用内存）
        int memoryBasedBatchSize = (int) ((availableMemory * 0.3) / estimatedRowSizeBytes);
        
        // 基于总数据量调整批处理大小
        int dataSizeBasedBatchSize;
        if (totalRows <= 10000) {
            dataSizeBasedBatchSize = 5000;
        } else if (totalRows <= 100000) {
            dataSizeBasedBatchSize = 20000;
        } else {
            dataSizeBasedBatchSize = DEFAULT_BATCH_SIZE;
        }
        
        // 综合计算，确保在合理范围内
        int optimalBatchSize = Math.max(DEFAULT_BATCH_SIZE / 2, 
                           Math.min(MAX_BATCH_SIZE, 
                           Math.max(memoryBasedBatchSize, dataSizeBasedBatchSize)));
        
        log.debug("批处理大小计算 - 基于内存: {}, 基于数据量: {}, 最终: {}", 
                memoryBasedBatchSize, dataSizeBasedBatchSize, optimalBatchSize);
        
        return optimalBatchSize;
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
            
            // 使用EasyExcel读取数据，并配置读取参数
            EasyExcel.read(file.getInputStream(), clazz, new DataReadListener<>(dataList))
                    .sheet()
                    .headRowNumber(1) // 明确指定表头行数
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
     * 内部使用的静态批处理大小计算方法，为了兼容旧代码保留
     */
    // 动态计算批次大小（兼容旧代码）
    public static int calculateBatchSize() {
        // 单条数据约1KB（根据实际模型估算）
        int singleDataSizeKB = 1;
        // 可用内存（MB）
        long freeMemoryMB = Runtime.getRuntime().freeMemory() / (1024 * 1024);
        // 用70%的可用内存来处理一批数据
        int batchSize = (int) (freeMemoryMB * 1024 / singleDataSizeKB * 0.7);
        // 限制批次大小在10000-50000之间（增大范围以适应大数据量）
        return Math.max(10000, Math.min(batchSize, 50000));
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
     * 增强版：添加行数限制检查，避免超出行数限制
     */
    private static class DataReadListener<T> extends AnalysisEventListener<T> {
        
        private final List<T> dataList;
        private static final int BATCH_COUNT = 5000; // 批处理大小
        private static final int EXCEL_MAX_ROWS = 1048576; // Excel 2007+ 最大行数
        private final List<T> batchList = new ArrayList<>(BATCH_COUNT);
        
        public DataReadListener(List<T> dataList) {
            this.dataList = dataList;
        }
        
        @Override
        public void invoke(T data, AnalysisContext context) {
            // 检查当前已读取数据量是否超过Excel最大行数限制
            if (dataList.size() + batchList.size() >= EXCEL_MAX_ROWS - 1) {
                throw new IllegalArgumentException("Excel文件包含过多数据行(" + (dataList.size() + batchList.size() + 1) + "行)，超出Excel最大限制(" + EXCEL_MAX_ROWS + "行)");
            }
            
            batchList.add(data);
            
            // 达到批处理数量就保存一次，减少内存占用
            if (batchList.size() >= BATCH_COUNT) {
                saveData();
            }
            
            // 每读取10000行记录一次进度
            if ((dataList.size() + batchList.size()) % 10000 == 0) {
                log.info("已读取数据行: {}", dataList.size() + batchList.size());
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
                // 使用同步块确保数据安全添加
                synchronized (dataList) {
                    dataList.addAll(batchList);
                }
                batchList.clear();
            }
        }
        
        @Override
        public void onException(Exception exception, AnalysisContext context) {
            log.error("读取Excel数据时发生异常：第{}行", context.readRowHolder().getRowIndex(), exception);
            // 继续抛出异常，确保调用者能感知到错误
            throw new RuntimeException("读取Excel数据失败", exception);
        }
    }
}