package com.example.excel.test;

import com.example.excel.entity.User;
import com.example.excel.excel.impl.ApachePoiExcelHandler;
import org.springframework.stereotype.Component;

/**
 * 泛型测试类，用于验证泛型类型的正确性
 */
@Component
public class GenericTest {

    // 这是一个简单的方法，用于验证泛型类型的正确性
    public ApachePoiExcelHandler<User> createPoiHandler() {
        return new ApachePoiExcelHandler<>();
    }

}