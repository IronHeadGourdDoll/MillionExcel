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

    @ExcelProperty("用户名")
    private String username;

    @ExcelProperty("姓名")
    private String name;

    @ExcelProperty("邮箱")
    private String email;

    @ExcelProperty("手机号")
    private String phone;

    @ExcelProperty("年龄")
    private Integer age;

    @ExcelProperty("创建时间")
    private LocalDateTime createTime;

    @ExcelProperty("更新时间")
    private LocalDateTime updateTime;

}