package com.example.excel.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import com.example.excel.entity.User;
import org.apache.ibatis.annotations.Mapper;

/**
 * 用户Mapper接口，继承MyBatis Plus的BaseMapper，提供基本的CRUD操作
 */
@Mapper
public interface UserMapper extends BaseMapper<User> {

}