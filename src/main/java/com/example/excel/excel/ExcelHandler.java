package com.example.excel.excel;

import com.example.excel.excel.impl.ApachePoiExcelHandler;
import com.example.excel.excel.impl.CsvExcelHandler;
import com.example.excel.excel.impl.EasyExcelHandler;
import org.springframework.web.multipart.MultipartFile;

import jakarta.servlet.http.HttpServletResponse;

import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.Future;
import java.util.concurrent.ConcurrentHashMap;
import java.util.Map;

/**
 * Excel处理接口，定义Excel导入导出的基本操作
 * 包含工厂方法和默认实现
 */
public interface ExcelHandler<T> {
    /**
     * Excel处理器类型枚举
     */
    enum HandlerType {
        EASY_EXCEL("easy"),
        CSV("csv"),
        APACHE_POI("poi");

        private final String code;

        HandlerType(String code) {
            this.code = code;
        }

        public String getCode() {
            return code;
        }
    }

    // 默认实现缓存
    static final Map<String, ExcelHandler<?>> handlerCache = new ConcurrentHashMap<>();
    static final EasyExcelHandler<?> DEFAULT_HANDLER = new EasyExcelHandler<>();

    /**
     * 获取Excel处理器实例
     * @param type 处理器类型
     * @param clazz 数据类型
     * @return Excel处理器实例
     */
    @SuppressWarnings("unchecked")
    static <T> ExcelHandler<T> getInstance(HandlerType type, Class<T> clazz) {
        String cacheKey = type.getCode() + ":" + clazz.getName();
        return (ExcelHandler<T>) handlerCache.computeIfAbsent(cacheKey, k -> {
            switch (type) {
                case CSV:
                    return new CsvExcelHandler<>();
                case APACHE_POI:
                    return new ApachePoiExcelHandler<>();
                case EASY_EXCEL:
                default:
                    return DEFAULT_HANDLER;
            }
        });
    }

    /**
     * 获取默认的Excel处理器实例
     * @param clazz 数据类型
     * @return 默认的Excel处理器实例
     */
    static <T> ExcelHandler<T> getDefaultInstance(Class<T> clazz) {
        return getInstance(HandlerType.EASY_EXCEL, clazz);
    }

    /**
     * 导出数据到Excel文件
     * @param dataList 数据列表
     * @param response HTTP响应对象
     * @param filename 文件名
     * @param clazz 数据类型
     */
    void export(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) throws UnsupportedEncodingException;

    /**
     * 从Excel文件导入数据
     * @param file 上传的Excel文件
     * @param clazz 数据类型
     * @return 导入的数据列表
     */
    List<T> importExcel(MultipartFile file, Class<T> clazz);

    /**
     * 异步导出数据到Excel文件
     * 默认实现使用CompletableFuture
     * @param dataList 数据列表
     * @param response HTTP响应对象
     * @param filename 文件名
     * @param clazz 数据类型
     * @return 导出结果的Future对象
     */
    default Future<Boolean> asyncExport(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) {
        return CompletableFuture.supplyAsync(() -> {
            try {
                export(dataList, response, filename, clazz);
                return true;
            } catch (Exception e) {
                return false;
            }
        });
    }

    /**
     * 异步从Excel文件导入数据
     * 默认实现使用CompletableFuture
     * @param file 上传的Excel文件
     * @param clazz 数据类型
     * @return 导入的数据列表的Future对象
     */
    default Future<List<T>> asyncImportExcel(MultipartFile file, Class<T> clazz) {
        return CompletableFuture.supplyAsync(() -> {
            try {
                return importExcel(file, clazz);
            } catch (Exception e) {
                return new ArrayList<>();
            }
        });
    }
}