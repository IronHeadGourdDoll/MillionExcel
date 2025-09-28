package com.example.excel.excel.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.example.excel.excel.ExcelHandler;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Async;
import org.springframework.scheduling.annotation.AsyncResult;
import org.springframework.stereotype.Component;
import org.springframework.util.StopWatch;
import org.springframework.web.multipart.MultipartFile;

import jakarta.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Future;

/**
 * 方案2：使用EasyExcel实现Excel导入导出
 * 特点：阿里巴巴开源的高性能Excel处理库，内存占用低，适合大数据量
 */
@Slf4j
public class EasyExcelHandler<T> implements ExcelHandler<T> {

    @Value("${excel.easyexcel.batch-size}")
    private int batchSize;

    @Value("${excel.easyexcel.thread-count}")
    private int threadCount;

    @Override
    public void export(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) {
        long startTime = System.currentTimeMillis();
        
        try {
            // 设置响应头
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition", 
                "attachment; filename=" + URLEncoder.encode(filename, StandardCharsets.UTF_8.toString()));
            
            // 使用EasyExcel写入数据
            EasyExcel.write(response.getOutputStream(), clazz)
                    .sheet("Sheet1")
                    .doWrite(dataList);
            
            long endTime = System.currentTimeMillis();
            log.info("EasyExcel导出完成，共导出{}条数据，耗时{}ms", dataList.size(), (endTime - startTime));
        } catch (Exception e) {
            log.error("EasyExcel导出Excel失败", e);
            throw new RuntimeException("导出Excel失败", e);
        }
    }

    @Override
    @Async
    public Future<Boolean> asyncExport(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) {
        try {
            this.export(dataList, response, filename, clazz);
            return new AsyncResult<>(true);
        } catch (Exception e) {
            log.error("异步EasyExcel导出Excel失败", e);
            return new AsyncResult<>(false);
        }
    }

    @Override
    public List<T> importExcel(MultipartFile file, Class<T> clazz) {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        List<T> dataList = new ArrayList<>();
        
        try (InputStream inputStream = file.getInputStream()) {
            log.info("开始EasyExcel导入文件：{}，大小：{}KB", file.getOriginalFilename(), file.getSize() / 1024);
            
            // 设置读取配置以优化性能
            EasyExcel.read(inputStream, clazz, new DataReadListener<>(dataList))
                    .autoCloseStream(false) // 我们自己控制流的关闭
                    .sheet()
                    .headRowNumber(1) // 表头行号，默认为0
                    .doRead();
            
            stopWatch.stop();
            log.info("EasyExcel导入完成，共导入{}条数据，耗时{}ms", dataList.size(), stopWatch.getTotalTimeMillis());
        } catch (Exception e) {
            log.error("EasyExcel导入Excel失败", e);
            throw new RuntimeException("导入Excel失败", e);
        }
        
        return dataList;
    }

    @Override
    @Async
    public Future<List<T>> asyncImportExcel(MultipartFile file, Class<T> clazz) {
        try {
            List<T> dataList = this.importExcel(file, clazz);
            return new AsyncResult<>(dataList);
        } catch (Exception e) {
            log.error("异步EasyExcel导入Excel失败", e);
            return new AsyncResult<>(new ArrayList<>());
        }
    }

    @Override
    public void exportByPage(int pageNum, int pageSize, HttpServletResponse response, String filename, Class<T> clazz) {
        long startTime = System.currentTimeMillis();
        
        try {
            // 设置响应头
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition", 
                "attachment; filename=" + URLEncoder.encode(filename, StandardCharsets.UTF_8.toString()));
            
            // 注意：由于移除了UserService依赖，此方法当前只创建空文件
            // 在实际应用中，这里应该有一个泛型的服务接口来获取相应类型的数据
            log.warn("EasyExcel分页导出：当前模式仅创建空文件，数据获取功能已临时禁用");
            
            // 创建一个空列表以避免NPE
            List<T> emptyList = new ArrayList<>();
            
            // 使用传入的clazz类型创建ExcelWriter
            EasyExcel.write(response.getOutputStream(), clazz).sheet("Sheet1").doWrite(emptyList);
            
            long endTime = System.currentTimeMillis();
            log.info("EasyExcel分页导出完成，耗时{}ms", (endTime - startTime));
        } catch (Exception e) {
            log.error("EasyExcel分页导出Excel失败", e);
            throw new RuntimeException("导出Excel失败", e);
        }
    }

    /**
     * 数据读取监听器，用于EasyExcel异步读取数据
     * 优化后的监听器支持更高效的数据处理
     */
    private static class DataReadListener<T> extends AnalysisEventListener<T> {
        
        private final List<T> dataList;
        private static final int BATCH_COUNT = 10000; // 批处理大小，根据内存情况调整
        private List<T> batchList = new ArrayList<>(BATCH_COUNT); // 预分配空间
        private int totalCount = 0;
        private final StopWatch stopWatch = new StopWatch();
        
        public DataReadListener(List<T> dataList) {
            this.dataList = dataList;
            stopWatch.start();
        }
        
        @Override
        public void invoke(T data, AnalysisContext context) {
            batchList.add(data);
            totalCount++;
            
            // 达到批处理数量就保存一次，减少内存占用
            if (batchList.size() >= BATCH_COUNT) {
                saveData();
                // 定期记录进度
                if (totalCount % 50000 == 0) {
                    log.info("已读取{}条数据", totalCount);
                }
            }
        }
        
        @Override
        public void doAfterAllAnalysed(AnalysisContext context) {
            // 保存最后一批数据
            saveData();
            stopWatch.stop();
            log.info("EasyExcel读取完成，共读取{}条数据，耗时{}ms", totalCount, stopWatch.getTotalTimeMillis());
        }
        
        private void saveData() {
            if (!batchList.isEmpty()) {
                dataList.addAll(batchList);
                batchList.clear(); // 清空批处理列表以释放内存
            }
        }
        
        @Override
        public void onException(Exception exception, AnalysisContext context) {
            log.error("读取Excel数据时发生异常：第{}行", context.readRowHolder().getRowIndex(), exception);
            // 可以选择继续处理下一行，或者抛出异常中断处理
        }
    }

}