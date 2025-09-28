package com.example.excel.excel;

import org.springframework.web.multipart.MultipartFile;

import jakarta.servlet.http.HttpServletResponse;
import java.util.List;
import java.util.concurrent.Future;

/**
 * Excel处理接口，定义Excel导入导出的基本操作
 */
public interface ExcelHandler<T> {

    /**
     * 导出数据到Excel文件
     * @param dataList 数据列表
     * @param response HTTP响应对象
     * @param filename 文件名
     * @param clazz 数据类型
     */
    void export(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz);

    /**
     * 异步导出数据到Excel文件
     * @param dataList 数据列表
     * @param response HTTP响应对象
     * @param filename 文件名
     * @param clazz 数据类型
     * @return 导出结果的Future对象
     */
    Future<Boolean> asyncExport(List<T> dataList, HttpServletResponse response, String filename, Class<T> clazz);

    /**
     * 从Excel文件导入数据
     * @param file 上传的Excel文件
     * @param clazz 数据类型
     * @return 导入的数据列表
     */
    List<T> importExcel(MultipartFile file, Class<T> clazz);

    /**
     * 异步从Excel文件导入数据
     * @param file 上传的Excel文件
     * @param clazz 数据类型
     * @return 导入的数据列表的Future对象
     */
    Future<List<T>> asyncImportExcel(MultipartFile file, Class<T> clazz);

    /**
     * 分页导出数据
     * @param pageNum 页码
     * @param pageSize 每页大小
     * @param response HTTP响应对象
     * @param filename 文件名
     * @param clazz 数据类型
     */
    void exportByPage(int pageNum, int pageSize, HttpServletResponse response, String filename, Class<T> clazz);

}