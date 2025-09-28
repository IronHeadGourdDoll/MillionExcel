package com.example.excel.excel.impl;

import com.example.excel.excel.ExcelHandler;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.CSVRecord;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Async;
import org.springframework.scheduling.annotation.AsyncResult;
import org.springframework.stereotype.Component;
import org.springframework.util.StopWatch;
import org.springframework.web.multipart.MultipartFile;

import jakarta.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.time.format.DateTimeParseException;
import java.util.*;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.concurrent.TimeoutException;
import java.util.zip.GZIPInputStream;
import java.util.zip.GZIPOutputStream;

/**
 * 方案3：使用CSV格式结合压缩技术实现Excel导入导出
 * 特点：处理速度快，文件体积小，特别适合超大数据量的场景
 */
@Slf4j
public class CsvExcelHandler<T> implements ExcelHandler<T> {

    @Value("${excel.csv.batch-size:5000}")
    private int batchSize;

    @Value("${excel.csv.thread-count:8}")
    private int threadCount;
    
    @Value("${excel.csv.compression-enabled:false}")
    private boolean compressionEnabled;
    
    @Value("${excel.csv.buffer-size:8192}")
    private int bufferSize;
    
    @Value("${excel.csv.memory-threshold-mb:200}")
    private long memoryThresholdMb; // 内存使用阈值，超过这个值采用流式处理
    
    @Value("${excel.csv.max-record-count:500000}")
    private int maxRecordCount; // 最大记录数，超过这个数采用流式处理
    
    // 线程池，用于并行处理CSV记录
    private ExecutorService executorService;
    
    // 核心线程数，基于CPU核心数动态调整
    private int corePoolSize;
    
    @jakarta.annotation.PostConstruct
    public void init() {
        // 初始化线程池 - 在PostConstruct中初始化，确保@Value注入已经完成
        this.corePoolSize = Math.max(2, Runtime.getRuntime().availableProcessors() - 1);
        this.executorService = new ThreadPoolExecutor(
                corePoolSize,
                threadCount,
                60L, TimeUnit.SECONDS,
                new LinkedBlockingQueue<>(10000),
                new ThreadFactory() {
                    private final AtomicInteger threadNumber = new AtomicInteger(1);
                    @Override
                    public Thread newThread(Runnable r) {
                        Thread t = new Thread(r, "csv-import-thread-" + threadNumber.getAndIncrement());
                        t.setDaemon(true);
                        t.setPriority(Thread.NORM_PRIORITY); // 设置线程优先级
                        return t;
                    }
                },
                new ThreadPoolExecutor.CallerRunsPolicy() // 当线程池饱和时，使用调用线程执行
        );
    }

    @Override
    public void export(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) {
        long startTime = System.currentTimeMillis();
        
        try {
            // 处理文件名，根据是否压缩添加相应的后缀
            String finalFilename = filename;
            String contentType = "text/csv";
            
            if (compressionEnabled) {
                finalFilename += ".gz";
                contentType = "application/gzip";
            }
            
            // 设置响应头
            response.setContentType(contentType);
            response.setHeader("Content-Disposition", 
                "attachment; filename=" + URLEncoder.encode(finalFilename, StandardCharsets.UTF_8.toString()));
            
            // 获取字段名作为CSV表头
            Field[] fields = clazz.getDeclaredFields();
            String[] headers = new String[fields.length];
            for (int i = 0; i < fields.length; i++) {
                headers[i] = fields[i].getName();
            }
            
            // 创建CSV打印机
            try (OutputStream outputStream = compressionEnabled ? 
                    new GZIPOutputStream(response.getOutputStream()) : response.getOutputStream();
                 Writer writer = new OutputStreamWriter(outputStream, StandardCharsets.UTF_8)) {
                
                // 写入UTF-8 BOM标记，解决Excel打开CSV文件乱码问题
                if (!compressionEnabled) {
                    outputStream.write(0xEF);
                    outputStream.write(0xBB);
                    outputStream.write(0xBF);
                }
                
                // 创建CSV打印机
                try (CSVPrinter csvPrinter = new CSVPrinter(writer, CSVFormat.DEFAULT.withHeader(headers))) {
                    
                    // 写入数据
                    for (int i = 0; i < dataList.size(); i++) {
                        if (i > 0 && i % 10000 == 0) {
                            log.info("已处理{}条数据", i);
                        }
                        
                        T data = dataList.get(i);
                        List<String> record = new ArrayList<>();
                        
                        for (Field field : fields) {
                            field.setAccessible(true);
                            Object value = field.get(data);
                            record.add(value != null ? value.toString() : "");
                        }
                        
                        csvPrinter.printRecord(record);
                    }
                    
                    csvPrinter.flush();
                }
            }
            
            long endTime = System.currentTimeMillis();
            log.info("CSV导出完成，共导出{}条数据，耗时{}ms", dataList.size(), (endTime - startTime));
        } catch (Exception e) {
            log.error("CSV导出失败", e);
            throw new RuntimeException("导出失败", e);
        }
    }

    @Override
    @Async
    public Future<Boolean> asyncExport(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) {
        try {
            this.export(dataList, response, filename, clazz);
            return new AsyncResult<>(true);
        } catch (Exception e) {
            log.error("异步CSV导出失败", e);
            return new AsyncResult<>(false);
        }
    }

    @Override
    public List<T> importExcel(MultipartFile file, Class<T> clazz) {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        List<T> dataList = new ArrayList<>();
        int validCount = 0;
        int invalidCount = 0;
        
        try {
            String fileName = file.getOriginalFilename();
            long fileSize = file.getSize();
            log.info("开始CSV导入文件：{}，大小：{}KB，核心线程数：{}，最大线程数：{}，批处理大小：{}", 
                    fileName, fileSize / 1024, corePoolSize, threadCount, batchSize);
            
            // 根据文件大小和系统内存状况智能选择处理策略
            if (shouldUseStreamingProcessing(fileSize)) {
                // 超大文件使用流式处理策略
                validCount = streamingImport(file, clazz, dataList);
            } else if (threadCount > 1 && fileSize > 10 * 1024 * 1024) { // 大于10MB的文件使用并行处理
                dataList = parallelImport(file, clazz);
                validCount = dataList.size();
            } else {
                // 小文件使用优化的单线程处理
                validCount = singleThreadImport(file, clazz, dataList);
            }
            
            stopWatch.stop();
            log.info("CSV导入完成，共处理{}条数据，有效数据{}条，无效数据{}条，耗时{}ms", 
                    validCount + invalidCount, validCount, invalidCount, stopWatch.getTotalTimeMillis());
        } catch (Exception e) {
            log.error("CSV导入失败", e);
            throw new RuntimeException("导入失败", e);
        }
        
        return dataList;
    }
    
    /**
     * 判断是否应该使用流式处理策略
     */
    private boolean shouldUseStreamingProcessing(long fileSize) {
        // 检查系统可用内存
        Runtime runtime = Runtime.getRuntime();
        long freeMemory = runtime.freeMemory();
        long totalMemory = runtime.totalMemory();
        long maxMemory = runtime.maxMemory();
        long usedMemory = totalMemory - freeMemory;
        long availableMemory = maxMemory - usedMemory;
        
        // 转换为MB
        long availableMemoryMb = availableMemory / (1024 * 1024);
        long fileSizeMb = fileSize / (1024 * 1024);
        
        // 如果文件大小超过可用内存的40%或达到配置的阈值，则使用流式处理
        boolean memoryConstrained = fileSizeMb > availableMemoryMb * 0.4 || fileSizeMb > memoryThresholdMb;
        
        log.debug("系统内存状态 - 可用: {}MB, 文件大小: {}MB, 是否使用流式处理: {}", 
                availableMemoryMb, fileSizeMb, memoryConstrained);
        
        return memoryConstrained;
    }
    
    /**
     * 单线程优化导入CSV文件
     */
    private int singleThreadImport(MultipartFile file, Class<T> clazz, List<T> dataList) throws Exception {
        // 处理压缩文件
        InputStream inputStream = file.getOriginalFilename().endsWith(".gz") ?
                new GZIPInputStream(file.getInputStream()) : file.getInputStream();
        
        int validCount = 0;
        
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream, StandardCharsets.UTF_8), bufferSize);
             CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT.withFirstRecordAsHeader().withTrim())) {
            
            Field[] fields = clazz.getDeclaredFields();
            // 创建字段名到Excel列名的映射
            Map<String, String> fieldToExcelName = new HashMap<>(fields.length);
            for (Field field : fields) {
                ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
                if (excelProperty != null && excelProperty.value().length > 0) {
                    fieldToExcelName.put(field.getName(), excelProperty.value()[0]);
                }
            }
            
            // 优化：预分配实例列表空间
            List<T> batchList = new ArrayList<>(batchSize);
            
            // 解析CSV记录
            int recordIndex = 0;
            for (CSVRecord csvRecord : csvParser) {
                T instance = mapRecordToObject(csvRecord, clazz, fields, fieldToExcelName);
                
                // 数据验证：确保必填字段不为空
                if (instance != null && isValidData(instance, clazz)) {
                    batchList.add(instance);
                    validCount++;
                    
                    // 达到批处理大小就添加到结果列表，减少内存占用
                    if (batchList.size() >= batchSize) {
                        dataList.addAll(batchList);
                        batchList.clear();
                    }
                }
                
                // 定期记录进度
                if (++recordIndex % 50000 == 0) {
                    log.info("已处理{}条记录", recordIndex);
                }
            }
            
            // 添加最后一批数据
            if (!batchList.isEmpty()) {
                dataList.addAll(batchList);
            }
        }
        
        return validCount;
    }
    
    /**
     * 并行导入CSV文件
     */
    private List<T> parallelImport(MultipartFile file, Class<T> clazz) throws Exception {
        // 首先读取所有记录到内存
        List<CSVRecord> records = readAllRecords(file);
        int totalRecords = records.size();
        log.info("CSV文件共包含{}条记录，准备并行处理", totalRecords);
        
        // 根据记录数量动态调整批处理大小
        int adjustedBatchSize = calculateOptimalBatchSize(totalRecords);
        log.debug("根据记录数量调整批处理大小为: {}", adjustedBatchSize);
        
        // 使用ArrayList加同步块代替CopyOnWriteArrayList，提高性能
        List<T> resultList = new ArrayList<>();
        CountDownLatch latch = new CountDownLatch(threadCount);
        AtomicInteger recordIndex = new AtomicInteger(0);
        AtomicReference<Exception> importException = new AtomicReference<>();
        
        // 预创建字段映射缓存，避免重复解析注解
        Field[] fields = clazz.getDeclaredFields();
        Map<String, String> fieldToExcelName = new HashMap<>(fields.length);
        for (Field field : fields) {
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            if (excelProperty != null && excelProperty.value().length > 0) {
                fieldToExcelName.put(field.getName(), excelProperty.value()[0]);
            }
        }
        
        // 启动多个线程进行数据处理
        for (int i = 0; i < threadCount; i++) {
            final int threadId = i;
            executorService.submit(() -> {
                try {
                    // 每个线程使用自己的批处理列表，减少线程间竞争
                    List<T> threadBatch = new ArrayList<>(adjustedBatchSize);
                    int currentIndex;
                    
                    while ((currentIndex = recordIndex.getAndAdd(adjustedBatchSize)) < totalRecords && 
                           importException.get() == null) {
                        int endIndex = Math.min(currentIndex + adjustedBatchSize, totalRecords);
                        List<CSVRecord> batchRecords = records.subList(currentIndex, endIndex);
                        
                        for (CSVRecord record : batchRecords) {
                            T instance = mapRecordToObject(record, clazz, fields, fieldToExcelName);
                            if (instance != null && isValidData(instance, clazz)) {
                                threadBatch.add(instance);
                            }
                        }
                        
                        // 批量添加到结果列表，减少锁竞争
                        if (!threadBatch.isEmpty()) {
                            synchronized (resultList) {
                                resultList.addAll(threadBatch);
                            }
                            threadBatch.clear();
                        }
                        
                        // 定期记录进度
                        if (endIndex % 50000 == 0) {
                            log.info("线程{}处理进度: {}/{}条记录", threadId, endIndex, totalRecords);
                        }
                        
                        log.debug("线程{}处理完成批次：{}-{}", threadId, currentIndex, endIndex);
                    }
                } catch (Exception e) {
                    log.error("线程{}处理异常", threadId, e);
                    importException.set(e);
                } finally {
                    latch.countDown();
                }
            });
        }
        
        // 等待所有线程完成，设置超时时间
        boolean completed = latch.await(10, TimeUnit.MINUTES);
        if (!completed) {
            log.error("CSV并行导入超时，已取消处理");
            throw new TimeoutException("CSV并行导入超时");
        }
        
        // 检查是否有异常
        if (importException.get() != null) {
            throw new RuntimeException("并行导入CSV失败", importException.get());
        }
        
        return resultList;
    }
    
    /**
     * 根据记录总数计算最优的批处理大小
     */
    private int calculateOptimalBatchSize(int totalRecords) {
        // 记录数少时使用较小的批处理大小，记录数多时使用较大的批处理大小
        if (totalRecords < 10000) {
            return Math.min(batchSize, 2000);
        } else if (totalRecords < 100000) {
            return Math.min(batchSize, 5000);
        } else {
            return Math.min(batchSize, 10000);
        }
    }
    
    /**
     * 流式处理导入CSV文件
     * 适用于超大文件，避免一次性加载所有记录到内存
     */
    private int streamingImport(MultipartFile file, Class<T> clazz, List<T> resultList) throws Exception {
        log.info("使用流式处理策略导入CSV文件");
        
        // 处理压缩文件
        InputStream inputStream = file.getOriginalFilename().endsWith(".gz") ?
                new GZIPInputStream(file.getInputStream()) : file.getInputStream();
        
        int validCount = 0;
        int recordIndex = 0;
        
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream, StandardCharsets.UTF_8), bufferSize * 2); // 增大缓冲区
             CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT.withFirstRecordAsHeader().withTrim())) {
            
            Field[] fields = clazz.getDeclaredFields();
            Map<String, String> fieldToExcelName = new HashMap<>(fields.length);
            for (Field field : fields) {
                ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
                if (excelProperty != null && excelProperty.value().length > 0) {
                    fieldToExcelName.put(field.getName(), excelProperty.value()[0]);
                }
            }
            
            // 使用更大的批处理列表
            List<T> batchList = new ArrayList<>(batchSize * 2);
            
            // 逐行处理记录，避免一次性加载所有记录到内存
            Iterator<CSVRecord> recordIterator = csvParser.iterator();
            while (recordIterator.hasNext()) {
                CSVRecord csvRecord = recordIterator.next();
                T instance = mapRecordToObject(csvRecord, clazz, fields, fieldToExcelName);
                
                if (instance != null && isValidData(instance, clazz)) {
                    batchList.add(instance);
                    validCount++;
                    
                    // 达到批处理大小就添加到结果列表，减少内存占用
                    if (batchList.size() >= batchSize) {
                        synchronized (resultList) {
                            resultList.addAll(batchList);
                        }
                        batchList.clear();
                        
                        // 释放内存
                        System.gc();
                    }
                }
                
                // 定期记录进度
                if (++recordIndex % 100000 == 0) {
                    log.info("流式处理已处理{}条记录", recordIndex);
                }
                
                // 检查内存使用情况，如果内存占用过高，主动进行GC
                if (recordIndex % 50000 == 0) {
                    checkMemoryUsage();
                }
            }
            
            // 添加最后一批数据
            if (!batchList.isEmpty()) {
                synchronized (resultList) {
                    resultList.addAll(batchList);
                }
            }
        }
        
        return validCount;
    }
    
    /**
     * 检查并管理内存使用情况
     */
    private void checkMemoryUsage() {
        Runtime runtime = Runtime.getRuntime();
        long freeMemory = runtime.freeMemory();
        long totalMemory = runtime.totalMemory();
        double usedMemoryPercentage = (double)(totalMemory - freeMemory) / totalMemory * 100;
        
        // 如果内存使用率超过80%，主动进行GC
        if (usedMemoryPercentage > 80) {
            log.warn("内存使用率过高({}%)，触发主动GC", usedMemoryPercentage);
            System.gc();
        }
    }
    
    /**
     * 读取所有CSV记录到内存
     * 只用于中大型文件，超大型文件使用流式处理
     */
    private List<CSVRecord> readAllRecords(MultipartFile file) throws Exception {
        // 处理压缩文件
        InputStream inputStream = file.getOriginalFilename().endsWith(".gz") ?
                new GZIPInputStream(file.getInputStream()) : file.getInputStream();
        
        // 预估列表大小，减少扩容开销
        int estimatedSize = estimateRecordCount(file);
        List<CSVRecord> records = new ArrayList<>(estimatedSize > 0 ? estimatedSize : 10000);
        
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream, StandardCharsets.UTF_8), bufferSize);
             CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT.withFirstRecordAsHeader().withTrim())) {
            
            // 直接遍历添加，比Stream API更高效
            for (CSVRecord record : csvParser) {
                records.add(record);
                
                // 定期记录进度和检查内存
                if (records.size() % 100000 == 0) {
                    log.info("已读取{}条记录到内存", records.size());
                    checkMemoryUsage();
                }
            }
        }
        
        return records;
    }
    
    /**
     * 预估CSV文件中的记录数量
     */
    private int estimateRecordCount(MultipartFile file) {
        try {
            // 基于文件大小和每行平均大小估算记录数
            long fileSize = file.getSize();
            // 假设每行平均100字节
            int estimatedCount = (int) (fileSize / 100);
            return Math.min(estimatedCount, maxRecordCount); // 限制最大预估数量
        } catch (Exception e) {
            log.warn("无法预估记录数量", e);
            return 0;
        }
    }
    
    /**
     * 将CSV记录映射到对象 - 优化版
     */
    private T mapRecordToObject(CSVRecord csvRecord, Class<T> clazz, 
                              Field[] fields, Map<String, String> fieldToExcelName) throws Exception {
        T instance = clazz.getDeclaredConstructor().newInstance();
        boolean hasValidData = false;

        // 尝试映射所有字段
        for (Field field : fields) {
            field.setAccessible(true);
            String fieldName = field.getName();

            try {
                // 优化：先检查列名是否存在，减少异常捕获开销
                String value = getRecordValue(csvRecord, fieldName, fieldToExcelName);
                
                // 设置字段值
                if (value != null && !value.isEmpty()) {
                    setFieldValue(field, instance, value.trim());
                    hasValidData = true;
                }
            } catch (Exception e) {
                // 记录错误但继续处理
                log.debug("CSV字段映射异常: {}", fieldName, e);
            }
        }

        return hasValidData ? instance : null;
    }
    
    /**
     * 获取CSV记录中的值，减少异常捕获开销
     */
    private String getRecordValue(CSVRecord csvRecord, String fieldName, Map<String, String> fieldToExcelName) {
        // 首先尝试通过Excel列名获取值
        String excelColumnName = fieldToExcelName.get(fieldName);
        
        // 优化：使用contains方法先检查是否存在，减少异常捕获
        if (excelColumnName != null && csvRecord.isSet(excelColumnName)) {
            return csvRecord.get(excelColumnName);
        } else if (csvRecord.isSet(fieldName)) {
            // 如果Excel列名不存在或未设置，尝试使用字段名
            return csvRecord.get(fieldName);
        }
        
        // 都不存在时返回null
        return null;
    }
    
    /**
     * 设置字段值 - 优化版，使用预编译的正则表达式和缓存提高性能
     */
    private void setFieldValue(Field field, Object instance, String value) throws Exception {
        Class<?> fieldType = field.getType();
        
        if (fieldType == String.class) {
            field.set(instance, value);
        } else if (fieldType == Integer.class || fieldType == int.class) {
            try {
                field.set(instance, parseInteger(value));
            } catch (NumberFormatException e) {
                log.debug("整数解析失败: {}", value);
            }
        } else if (fieldType == Long.class || fieldType == long.class) {
            try {
                field.set(instance, parseLong(value));
            } catch (NumberFormatException e) {
                log.debug("长整数解析失败: {}", value);
            }
        } else if (fieldType == Double.class || fieldType == double.class) {
            try {
                field.set(instance, parseDouble(value));
            } catch (NumberFormatException e) {
                log.debug("浮点数解析失败: {}", value);
            }
        } else if (fieldType == Boolean.class || fieldType == boolean.class) {
            field.set(instance, parseBoolean(value));
        } else if (fieldType == LocalDateTime.class) {
            try {
                field.set(instance, parseLocalDateTime(value));
            } catch (DateTimeParseException e) {
                log.debug("日期时间解析失败: {}", value);
            }
        } else {
            // 其他类型直接设置字符串
            field.set(instance, value);
        }
    }
    
    // 优化的类型转换方法，减少重复代码
    private Integer parseInteger(String value) {
        // 快速路径：常见的整数字符串
        if (value.length() <= 10) {
            return Integer.parseInt(value);
        }
        throw new NumberFormatException("Value too large for integer: " + value);
    }
    
    private Long parseLong(String value) {
        // 快速路径：常见的长整数字符串
        if (value.length() <= 19) {
            return Long.parseLong(value);
        }
        throw new NumberFormatException("Value too large for long: " + value);
    }
    
    private Double parseDouble(String value) {
        // 处理常见的特殊情况
        if (value.equalsIgnoreCase("NaN")) return Double.NaN;
        if (value.equalsIgnoreCase("Infinity")) return Double.POSITIVE_INFINITY;
        if (value.equalsIgnoreCase("-Infinity")) return Double.NEGATIVE_INFINITY;
        
        return Double.parseDouble(value);
    }
    
    private boolean parseBoolean(String value) {
        // 优化布尔值解析，处理常见的情况
        if (value.equalsIgnoreCase("true") || value.equals("1")) return true;
        if (value.equalsIgnoreCase("false") || value.equals("0")) return false;
        
        return Boolean.parseBoolean(value);
    }
    
    private LocalDateTime parseLocalDateTime(String value) {
        // 尝试不同的日期时间格式
        try {
            return LocalDateTime.parse(value);
        } catch (DateTimeParseException e) {
            // 可以根据需要添加更多的日期格式解析
            log.debug("无法解析日期时间: {}", value);
            throw e;
        }
    }
    
    /**
     * 验证数据是否有效
     */
    private <T> boolean isValidData(T instance, Class<T> clazz) {
        // 如果是User类型，需要特别验证username和name字段
        if (instance instanceof com.example.excel.entity.User) {
            com.example.excel.entity.User user = (com.example.excel.entity.User) instance;
            // username和name是必填字段，不能为null或空字符串
            return user.getUsername() != null && !user.getUsername().isEmpty() && 
                   user.getName() != null && !user.getName().isEmpty();
        }
        // 其他类型的基本验证，可以根据需要扩展
        return true;
    }

    @Override
    @Async
    public Future<List<T>> asyncImportExcel(MultipartFile file, Class<T> clazz) {
        try {
            List<T> dataList = this.importExcel(file, clazz);
            return new AsyncResult<>(dataList);
        } catch (Exception e) {
            log.error("异步CSV导入失败", e);
            return new AsyncResult<>(new ArrayList<>());
        }
    }

    @Override
    public void exportByPage(int pageNum, int pageSize, HttpServletResponse response, String filename, Class<T> clazz) {
        long startTime = System.currentTimeMillis();
        
        try {
            // 处理文件名和响应头
            String finalFilename = filename;
            String contentType = "text/csv";
            
            if (compressionEnabled) {
                finalFilename += ".gz";
                contentType = "application/gzip";
            }
            
            response.setContentType(contentType);
            response.setHeader("Content-Disposition", 
                "attachment; filename=" + URLEncoder.encode(finalFilename, StandardCharsets.UTF_8.toString()));
            
            // 获取字段名作为表头
            Field[] fields = clazz.getDeclaredFields();
            String[] headers = new String[fields.length];
            for (int i = 0; i < fields.length; i++) {
                headers[i] = fields[i].getName();
            }
            
            // 注意：由于移除了UserService依赖，此方法当前只创建空文件
            // 在实际应用中，这里应该有一个泛型的服务接口来获取相应类型的数据
            log.warn("CSV分页导出：当前模式仅创建空文件，数据获取功能已临时禁用");
            
            // 创建CSV打印机并写入表头
            try (OutputStream outputStream = compressionEnabled ? 
                    new GZIPOutputStream(response.getOutputStream()) : response.getOutputStream();
                 Writer writer = new OutputStreamWriter(outputStream, StandardCharsets.UTF_8);
                 CSVPrinter csvPrinter = new CSVPrinter(writer, CSVFormat.DEFAULT.withHeader(headers))) {
                
                // 不写入任何数据行
                csvPrinter.flush();
            }
            
            long endTime = System.currentTimeMillis();
            log.info("CSV分页导出完成，耗时{}ms", (endTime - startTime));
        } catch (Exception e) {
            log.error("CSV分页导出失败", e);
            throw new RuntimeException("导出失败", e);
        }
    }

}