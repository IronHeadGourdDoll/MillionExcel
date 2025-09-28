package com.example.excel.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.transaction.annotation.EnableTransactionManagement;
import org.springframework.transaction.PlatformTransactionManager;
import javax.sql.DataSource;

/**
 * 事务配置类，启用Spring的事务管理功能
 */
@Configuration
@EnableTransactionManagement
public class TransactionConfig {

    /**
     * 配置事务管理器
     * @param dataSource 数据源
     * @return 事务管理器实例
     */
    @Bean
    public PlatformTransactionManager transactionManager(DataSource dataSource) {
        // 使用Spring的DataSourceTransactionManager
        org.springframework.jdbc.datasource.DataSourceTransactionManager transactionManager = 
                new org.springframework.jdbc.datasource.DataSourceTransactionManager();
        transactionManager.setDataSource(dataSource);
        return transactionManager;
    }
}