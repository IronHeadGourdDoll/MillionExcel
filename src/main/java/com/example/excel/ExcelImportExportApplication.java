package com.example.excel;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableAsync;

/**
 * Excel导入导出应用主类
 */
@SpringBootApplication
@EnableAsync
@MapperScan("com.example.excel.mapper")
public class ExcelImportExportApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExcelImportExportApplication.class, args);
    }

}