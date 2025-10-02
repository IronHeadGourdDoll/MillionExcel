package com.example.excel.controller;

import cn.hutool.core.date.StopWatch;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.example.excel.entity.User;
import com.example.excel.excel.ExcelHandler;
import com.example.excel.service.UserService;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.WriteListener;
import jakarta.servlet.http.HttpServletResponseWrapper;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.*;
import java.util.concurrent.ConcurrentHashMap;
import org.springframework.scheduling.annotation.Scheduled;

/**
 * Excel导入导出控制器，提供REST API接口
 */
@RestController
@RequestMapping("/api/excel")
@Slf4j
public class ExcelController {

    @Autowired
    private UserService userService;

    @Autowired
    private ExcelHandler excelHandler;

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

            // 获取用户数据
            List<User> users = userService.list();

            // 获取对应的Excel处理器
            ExcelHandler<User> excelHandler = ExcelHandler.getInstance(
                "easyexcel".equals(type) ? ExcelHandler.HandlerType.EASY_EXCEL :
                "csv".equals(type) ? ExcelHandler.HandlerType.CSV :
                ExcelHandler.HandlerType.APACHE_POI, 
                User.class
            );

            // 根据类型设置文件名后缀和内容类型
            String fileExtension = "xlsx";
            String contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            if ("csv".equals(type)) {
                fileExtension = "csv";
                contentType = "text/csv;charset=UTF-8";
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
            excelHandler.export(users, response, fileName, User.class);
            
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
     */
    // 临时文件跟踪器
    private static final Map<String, Long> fileExpiryMap = new ConcurrentHashMap<>();
    private static final long FILE_EXPIRY_TIME = 3600_000; // 1小时过期
    
    @GetMapping("/async-export")
    public String asyncExportExcel(@RequestParam(value = "type", defaultValue = "easyexcel") String type) {
        // 创建临时文件名
        String fileExtension = "csv".equals(type) ? "csv" : "xlsx";
        String fileName = "export_" + System.currentTimeMillis() + "." + fileExtension;
        String filePath = System.getProperty("java.io.tmpdir") + fileName;
        
        // 启动异步导出任务
        CompletableFuture.runAsync(() -> {
            try {
                // 获取用户数据
                List<User> users = userService.list();
                
                // 导出到临时文件
                try (OutputStream outputStream = new FileOutputStream(filePath)) {
                    ExcelWriter excelWriter = EasyExcel.write(outputStream, User.class)
                        .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                        .autoTrim(true)
                        .build();
                    
                    WriteSheet writeSheet = EasyExcel.writerSheet("Sheet1").build();
                    excelWriter.write(users, writeSheet);
                    excelWriter.finish();
                    
                    // 记录文件过期时间
                    fileExpiryMap.put(fileName, System.currentTimeMillis() + FILE_EXPIRY_TIME);
                    log.info("异步导出成功，文件已保存: {}", filePath);
                }
            } catch (Exception e) {
                log.error("异步导出失败", e);
                try {
                    Files.deleteIfExists(Paths.get(filePath));
                } catch (IOException ioException) {
                    log.error("删除临时文件失败", ioException);
                }
            }
        });
        
        // 立即返回下载URL
        return "/api/excel/download/" + fileName;
    }

    
    @GetMapping("/download/{fileName}")
    public void downloadExportFile(@PathVariable String fileName, HttpServletResponse response) {
        String filePath = System.getProperty("java.io.tmpdir") + fileName;
        
        try {
            Path path = Paths.get(filePath);
            if (!Files.exists(path)) {
                response.sendError(HttpServletResponse.SC_NOT_FOUND, "文件不存在或已过期");
                return;
            }
            
            // 设置响应头
            response.setContentType(Files.probeContentType(path));
            response.setHeader("Content-Disposition", "attachment; filename=" + fileName);
            response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
            response.setHeader("Pragma", "no-cache");
            response.setHeader("Expires", "0");
            
            // 传输文件
            Files.copy(path, response.getOutputStream());
            response.flushBuffer();
            
            // 下载完成后删除文件
            Files.deleteIfExists(path);
            fileExpiryMap.remove(fileName);
            log.info("文件下载完成并已删除: {}", filePath);
        } catch (IOException e) {
            log.error("文件下载失败", e);
            try {
                response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR, "文件下载失败");
            } catch (IOException ioException) {
                log.error("发送错误响应失败", ioException);
            }
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
            ExcelHandler<User> excelHandler = ExcelHandler.getInstance(
                "easyexcel".equals(type) ? ExcelHandler.HandlerType.EASY_EXCEL :
                "csv".equals(type) ? ExcelHandler.HandlerType.CSV :
                ExcelHandler.HandlerType.APACHE_POI, 
                User.class
            );

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
            excelHandler.asyncImportExcel(file, User.class);

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
            // 获取对应的Excel处理器
            ExcelHandler<User> excelHandler = ExcelHandler.getInstance(
                "easyexcel".equals(type) ? ExcelHandler.HandlerType.EASY_EXCEL :
                "csv".equals(type) ? ExcelHandler.HandlerType.CSV :
                ExcelHandler.HandlerType.APACHE_POI, 
                User.class
            );
            
            // 获取数据 - 优化性能
            StopWatch stopWatch = new StopWatch();
            stopWatch.start("query-data");
            List<User> users = userService.list();
            stopWatch.stop();
            log.debug("查询数据耗时: {}ms", stopWatch.getLastTaskTimeMillis());
            
            // 设置响应头
            String fileExtension = "csv".equals(type) ? "csv" : "xlsx";
            String contentType = "csv".equals(type) ? "text/csv;charset=UTF-8" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            
            response.setContentType(contentType);
            response.setCharacterEncoding(StandardCharsets.UTF_8.toString());
            response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
            response.setHeader("Pragma", "no-cache");
            response.setHeader("Expires", "0");
            
            // 执行导出
            excelHandler.export(users, response, "用户数据分页导出" + System.currentTimeMillis() + "." + fileExtension, User.class);
            
            log.info("分页导出完成，页码：1，每页大小：{}", pageSize);
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
     * 创建模拟用户数据
     *
     * @param count 用户数量
     * @return 模拟用户列表
     */
    // 定时清理过期文件
    @Scheduled(fixedRate = 3600_000) // 每小时清理一次
    public void cleanExpiredFiles() {
        long currentTime = System.currentTimeMillis();
        fileExpiryMap.entrySet().removeIf(entry -> {
            if (entry.getValue() < currentTime) {
                try {
                    Files.deleteIfExists(Paths.get(System.getProperty("java.io.tmpdir") + entry.getKey()));
                    log.info("已清理过期文件: {}", entry.getKey());
                    return true;
                } catch (IOException e) {
                    log.error("清理过期文件失败: {}", entry.getKey(), e);
                    return false;
                }
            }
            return false;
        });
    }

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