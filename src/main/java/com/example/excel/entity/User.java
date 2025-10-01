package com.example.excel.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import com.baomidou.mybatisplus.annotation.TableName;
import lombok.Data;
import java.time.LocalDateTime;

/**
 * 用户实体类，用于Excel导入导出测试
 */
@Data
@TableName("user")
public class User {

    @ExcelProperty("ID")
    private Long id;

    @ExcelProperty("username")
    private String username;

    @ExcelProperty("name")
    private String name;

    @ExcelProperty("email")
    private String email;

    @ExcelProperty("phone")
    private String phone;

    @ExcelProperty("age")
    private Integer age;

    @ExcelProperty("createTime")
    private LocalDateTime createTime;

    @ExcelProperty("updateTime")
    private LocalDateTime updateTime;

}