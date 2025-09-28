package com.example.excel.excel.impl;

import com.example.excel.excel.ExcelHandler;
import org.springframework.stereotype.Component;

import jakarta.servlet.http.HttpServletResponse;
import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.Future;

/**
 * 简单的Excel处理器实现
 */
@Component
public class SimpleExcelHandler<T> implements ExcelHandler<T> {

    @Override
    public void export(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) {
        try {
            // response.setContentType("application/vnd.ms-excel");
            response.setHeader("Content-Disposition", "attachment; filename=" + filename);
            response.getWriter().write("测试Excel数据");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public Future<Boolean> asyncExport(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz) {
        return CompletableFuture.completedFuture(true);
    }

    @Override
    public List<T> importExcel(MultipartFile file, Class<T> clazz) {
        return new ArrayList<>();
    }

    @Override
    public Future<List<T>> asyncImportExcel(MultipartFile file, Class<T> clazz) {
        return CompletableFuture.completedFuture(new ArrayList<>());
    }

    @Override
    public void exportByPage(int pageNum, int pageSize, HttpServletResponse response, String filename, Class<T> clazz) {
        // 简化实现
    }
}