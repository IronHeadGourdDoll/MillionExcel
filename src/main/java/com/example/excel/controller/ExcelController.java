package com.example.excel.controller;

import com.example.excel.entity.User;
import com.example.excel.excel.ExcelHandler;
import com.example.excel.excel.ExcelHandlerFactory;
import com.example.excel.excel.impl.UserExcelHandler;
import com.example.excel.service.UserService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.Future;

/**
 * Excel导入导出控制器，提供REST API接口
 */
@RestController
@RequestMapping("/api/excel")
@Slf4j
public class ExcelController {

    @Autowired
    private ExcelHandlerFactory excelHandlerFactory;

    @Autowired
    private UserService userService;

    @Autowired
    private UserExcelHandler userExcelHandler;

    /**
     * 导出数据到Excel - 统一接口
     *
     * @param type     导出类型：poi, easyexcel, csv
     * @param response HTTP响应
     */
    @GetMapping({
        "/export",
        "/export/apache-poi",
        "/export/easy-excel",
        "/export/csv"
    })
    public void exportExcel(@RequestParam(value = "type", required = false) String type,
                            HttpServletResponse response,
                            HttpServletRequest request) {
        long startTime = System.currentTimeMillis();

        try {
            // 根据路径确定导出类型，如果参数中没有提供
            if (type == null) {
                // 获取请求路径
                String requestUri = request.getRequestURI();
                if (requestUri.contains("/apache-poi")) {
                    type = "poi";
                } else if (requestUri.contains("/easy-excel")) {
                    type = "easyexcel";
                } else if (requestUri.contains("/csv")) {
                    type = "csv";
                } else {
                    // 默认使用easyexcel
                    type = "easyexcel";
                }
            }

            // 创建模拟数据用于测试 - 增加数据量到1000条
            List<User> mockUsers = createMockUsers(1000);

            // 获取对应的Excel处理器
            ExcelHandler<User> excelHandler = excelHandlerFactory.getExcelHandler(type, User.class);

            // 根据类型设置文件名后缀和内容类型
            String fileExtension = "xlsx";
            String contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            
            if ("csv".equals(type)) {
                fileExtension = "csv";
                contentType = "text/csv; charset=UTF-8";
            } else if ("poi".equals(type)) {
                // Apache POI默认生成xlsx格式
                fileExtension = "xlsx";
                contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            } else if ("easyexcel".equals(type)) {
                // EasyExcel生成xlsx格式
                fileExtension = "xlsx";
                contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            }

            // 设置响应头
            response.setContentType(contentType);
            response.setCharacterEncoding(StandardCharsets.UTF_8.toString());
            // 确保响应头不会被缓存
            response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
            response.setHeader("Pragma", "no-cache");
            response.setHeader("Expires", "0");
            // 防止浏览器嗅探内容类型
            response.setHeader("X-Content-Type-Options", "nosniff");

            // 生成文件名（包含时间戳，避免缓存）
            String fileName = "用户数据_" + type + "_" + System.currentTimeMillis() + "." + fileExtension;
            
            // 执行导出
            try {
                excelHandler.export(mockUsers, response, fileName, User.class);
                // 确保响应完全刷新
                response.flushBuffer();
            } catch (Exception e) {
                log.error("导出执行失败", e);
                throw e;
            }

            long endTime = System.currentTimeMillis();
            log.info("{}导出完成，耗时{}ms", type.toUpperCase(), (endTime - startTime));
        } catch (Exception e) {
            log.error("导出Excel失败", e);
            try {
                response.getWriter().write("导出Excel失败：" + e.getMessage());
            } catch (IOException ioException) {
                log.error("写入响应失败", ioException);
            }
        }
    }

    /**
     * 异步导出数据到Excel
     *
     * @param type     导出类型：poi, easyexcel, csv
     * @param response HTTP响应
     */
    @GetMapping("/async-export")
    public void asyncExportExcel(@RequestParam(value = "type", defaultValue = "easyexcel") String type,
                                 HttpServletResponse response) {
        try {
            // 注意：由于移除了UserService依赖，创建模拟数据用于测试
            List<User> mockUsers = createMockUsers(10);

            // 根据类型设置文件名后缀和内容类型
            String fileExtension = "xlsx";
            String contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            
            if ("csv".equals(type)) {
                fileExtension = "csv";
                contentType = "text/csv; charset=UTF-8";
            }
            
            // 设置响应头
            response.setContentType(contentType);
            response.setCharacterEncoding(StandardCharsets.UTF_8.toString());
            // 确保响应头不会被缓存
            response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
            response.setHeader("Pragma", "no-cache");
            response.setHeader("Expires", "0");

            // 执行异步导出
            Future<Boolean> future = userExcelHandler.asyncExport(mockUsers, response, "用户数据_异步_" + System.currentTimeMillis() + "." + fileExtension, User.class);

            // 轮询等待完成（在实际生产环境中，可能需要返回任务ID，让客户端轮询任务状态）
            CompletableFuture.supplyAsync(() -> {
                try {
                    return future.get();
                } catch (InterruptedException | ExecutionException e) {
                    log.error("异步导出Excel异常", e);
                    return false;
                }
            }).thenAccept(result -> {
                if (result) {
                    log.info("异步导出Excel成功");
                } else {
                    log.error("异步导出Excel失败");
                }
            });
        } catch (Exception e) {
            log.error("异步导出Excel失败", e);
            try {
                response.getWriter().write("异步导出Excel失败：" + e.getMessage());
            } catch (IOException ioException) {
                log.error("写入响应失败", ioException);
            }
        }
    }

    /**
     * 从Excel导入数据
     *
     * @param type 导入类型：poi, easyexcel, csv
     * @param file 上传的Excel文件
     * @return 导入结果
     */
    @PostMapping("/import")
    public String importExcel(@RequestParam(value = "type", defaultValue = "easyexcel") String type,
                              @RequestParam("file") MultipartFile file) {
        long startTime = System.currentTimeMillis();

        try {
            if (file.isEmpty()) {
                return "上传的文件为空";
            }

            // 获取对应的Excel处理器
            ExcelHandler<User> excelHandler = excelHandlerFactory.getExcelHandler(type, User.class);

            // 执行导入
            List<User> users = excelHandler.importExcel(file, User.class);

            int savedCount = userService.batchSave(users);

            long endTime = System.currentTimeMillis();
            log.info("导入解析完成，共解析{}条数据，成功保存{}条数据，耗时{}ms", users.size(), savedCount, (endTime - startTime));

            return "导入解析成功，共解析" + users.size() + "条数据，成功保存" + savedCount + "条数据，耗时" + (endTime - startTime) + "ms";
        } catch (Exception e) {
            log.error("导入Excel失败", e);
            return "导入Excel失败：" + e.getMessage();
        }
    }

    /**
     * 异步从Excel导入数据
     *
     * @param type 导入类型：poi, easyexcel, csv
     * @param file 上传的Excel文件
     * @return 导入结果
     */
    @PostMapping("/async-import")
    public String asyncImportExcel(@RequestParam(value = "type", defaultValue = "easyexcel") String type,
                                   @RequestParam("file") MultipartFile file) {
        try {
            if (file.isEmpty()) {
                return "上传的文件为空";
            }

            // 执行异步导入
            Future<List<User>> future = userExcelHandler.asyncImportExcel(file, User.class);

            // 异步处理导入结果
            CompletableFuture.supplyAsync(() -> {
                try {
                    return future.get();
                } catch (InterruptedException | ExecutionException e) {
                    log.error("异步导入Excel异常", e);
                    return null;
                }
            }).thenAccept(users -> {
                if (users != null && !users.isEmpty()) {
                    try {
                        userService.batchSave(users);
                        log.info("异步导入解析完成，共解析并保存{}条数据", users.size());
                    } catch (ExecutionException | InterruptedException e) {
                        log.error("异步保存用户数据异常", e);
                    }
                }
            });

            return "异步导入任务已提交，请稍后查看结果";
        } catch (Exception e) {
            log.error("异步导入Excel失败", e);
            return "异步导入Excel失败：" + e.getMessage();
        }
    }

    /**
     * 分页导出数据
     *
     * @param type     导出类型：poi, easyexcel, csv
     * @param pageSize 每页大小
     * @param response HTTP响应
     */
    @GetMapping("/export-by-page")
    public void exportByPage(@RequestParam(value = "type", defaultValue = "easyexcel") String type,
                             @RequestParam(value = "pageSize", defaultValue = "10000") int pageSize,
                             HttpServletResponse response) {
        try {
            // 执行分页导出
            userExcelHandler.exportByPage(1, pageSize, response, "用户数据分页导出" + System.currentTimeMillis() + ".xlsx", User.class);
        } catch (Exception e) {
            log.error("分页导出Excel失败", e);
            try {
                response.getWriter().write("分页导出Excel失败：" + e.getMessage());
            } catch (IOException ioException) {
                log.error("写入响应失败", ioException);
            }
        }
    }

    /**
     * 生成测试数据 - 注意：当前已禁用（无数据库连接）
     *
     * @param count 数据量
     * @return 生成结果
     */
    @GetMapping("/generate-test-data")
    public String generateTestData(@RequestParam(value = "count", defaultValue = "100000") int count) {
        log.warn("生成测试数据：功能已禁用（数据库服务未连接）");
        return "生成测试数据功能暂时不可用（数据库服务未连接）";
    }

    /**
     * 清空测试数据 - 注意：当前已禁用（无数据库连接）
     *
     * @return 清空结果
     */
    @GetMapping("/clear-test-data")
    public String clearTestData() {
        log.warn("清空测试数据：功能已禁用（数据库服务未连接）");
        return "清空测试数据功能暂时不可用（数据库服务未连接）";
    }

    /**
     * 下载导入模板
     */
    @GetMapping("/template")
    public void downloadTemplate(HttpServletResponse response) {
        try {
            // 设置响应头
            response.setContentType("text/csv");
            response.setHeader("Content-Disposition", "attachment; filename=user_import_template.csv");

            // 读取模板文件
            ClassPathResource resource = new ClassPathResource("templates/user_import_template.csv");
            try (InputStream inputStream = resource.getInputStream();
                 OutputStream outputStream = response.getOutputStream()) {
                byte[] buffer = new byte[1024];
                int bytesRead;
                while ((bytesRead = inputStream.read(buffer)) != -1) {
                    outputStream.write(buffer, 0, bytesRead);
                }
                outputStream.flush();
            }
        } catch (Exception e) {
            log.error("下载模板失败", e);
        }
    }

    /**
     * 下载修复后的导入模板（无#注释符号，正确的UTF-8编码）
     */
    @GetMapping("/corrected-template")
    public void downloadCorrectedTemplate(HttpServletResponse response) {
        try {
            // 设置响应头
            response.setContentType("text/csv");
            response.setHeader("Content-Disposition", "attachment; filename=user_import_template_corrected.csv");

            // 读取修复后的模板文件
            ClassPathResource resource = new ClassPathResource("templates/user_import_template_corrected.csv");
            try (InputStream inputStream = resource.getInputStream();
                 OutputStream outputStream = response.getOutputStream()) {
                byte[] buffer = new byte[1024];
                int bytesRead;
                while ((bytesRead = inputStream.read(buffer)) != -1) {
                    outputStream.write(buffer, 0, bytesRead);
                }
                outputStream.flush();
            }
        } catch (Exception e) {
            log.error("下载修复后模板失败", e);
        }
    }

    /**
     * 创建模拟用户数据
     *
     * @param count 用户数量
     * @return 模拟用户列表
     */
    private List<User> createMockUsers(int count) {
        List<User> users = new ArrayList<>();
        for (int i = 0; i < count; i++) {
            User user = new User();
            user.setId((long) i);
            user.setUsername("user" + i);
            user.setName("测试用户" + i);
            user.setEmail("user" + i + "@example.com");
            user.setAge(20 + i % 30);
            user.setPhone("1380013800" + (i % 10));
            LocalDateTime now = LocalDateTime.now();
            user.setCreateTime(now);
            user.setUpdateTime(now);
            users.add(user);
        }
        return users;
    }

}