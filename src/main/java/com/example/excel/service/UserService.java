package com.example.excel.service;

import com.baomidou.mybatisplus.extension.service.IService;
import com.example.excel.entity.User;

import java.util.List;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.Future;

/**
 * 用户Service接口，继承MyBatis Plus的IService，提供更多的业务操作方法
 */
public interface UserService extends IService<User> {

    /**
     * 批量保存用户数据
     * @param users 用户列表
     * @return 保存的记录数
     */
    int batchSave(List<User> users) throws ExecutionException, InterruptedException;

    /**
     * 异步批量保存用户数据
     * @param users 用户列表
     * @return 保存的记录数Future对象
     */
    Future<Integer> asyncBatchSave(List<User> users);

    /**
     * 分页查询用户数据
     * @param pageNum 页码
     * @param pageSize 每页大小
     * @return 用户列表
     */
    List<User> selectPage(int pageNum, int pageSize);

    /**
     * 生成测试数据
     * @param count 数据量
     */
    void generateTestData(int count);

}