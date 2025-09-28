package com.example.excel.excel;

import com.example.excel.excel.impl.ApachePoiExcelHandler;
import com.example.excel.excel.impl.CsvExcelHandler;
import com.example.excel.excel.impl.EasyExcelHandler;
import com.example.excel.excel.impl.UserExcelHandler;
import com.example.excel.entity.User;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Excel处理器工厂
 */
@Component
public class ExcelHandlerFactory {

    @Autowired
    private UserExcelHandler userExcelHandler;

    /**
     * 缓存已创建的处理器实例
     */
    private final Map<String, ExcelHandler<?>> handlerCache = new ConcurrentHashMap<>();

    /**
     * 获取Excel处理器
     */
    @SuppressWarnings("unchecked")
    public <T> ExcelHandler<T> getExcelHandler(String type, Class<T> clazz) {
        // 根据类型和目标类获取处理器
        String cacheKey = type + ":" + clazz.getName();
        return (ExcelHandler<T>) handlerCache.computeIfAbsent(cacheKey, k -> {
            // 对于User类型，使用注入的userExcelHandler
            if (User.class.equals(clazz)) {
                return userExcelHandler;
            }
            // 对于其他类型，根据type参数选择不同的处理器
            if ("poi".equals(type)) {
                return new ApachePoiExcelHandler<>();
            } else if ("csv".equals(type)) {
                return new CsvExcelHandler<>();
            } else {
                // 默认为EasyExcel处理器
                return new EasyExcelHandler<>();
            }
        });
    }
}