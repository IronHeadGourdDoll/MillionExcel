package com.example.excel.service.impl;

import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.example.excel.entity.User;
import com.example.excel.mapper.UserMapper;
import com.example.excel.service.UserService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Async;
import org.springframework.scheduling.annotation.AsyncResult;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.util.StopWatch;
import jakarta.annotation.PostConstruct;

import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Random;
import java.util.concurrent.*;
import java.util.concurrent.CompletionException;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

/**
 * 用户Service实现类，处理用户相关的业务逻辑
 */
@Service
@Slf4j
public class UserServiceImpl extends ServiceImpl<UserMapper, User> implements UserService {

    @Autowired
    private UserMapper userMapper;

    @Value("${excel.common.concurrent-save-enabled}")
    private boolean concurrentSaveEnabled;

    @Value("${excel.common.max-concurrent-saves}")
    private int maxConcurrentSaves;

    @Value("${excel.common.max-queue-size}")
    private int maxQueueSize;

    @Value("${excel.common.async-executor-prefix}")
    private String executorPrefix;

    // 线程池配置
    private ExecutorService saveThreadPool;

    private final Random random = new Random();
    private final int BATCH_SIZE = 10000; // 数据库批次大小，增加批次大小减少SQL执行次数
    private final int PARALLEL_BATCH_SIZE = 30000; // 并行处理时每个线程处理的数据量，增加批次大小减少线程开销
    
    // 使用@PostConstruct确保在所有属性注入后初始化线程池
    @PostConstruct
    public void init() {
        // 初始化线程池
        int corePoolSize = Runtime.getRuntime().availableProcessors();
        int defaultMaxPoolSize = Runtime.getRuntime().availableProcessors() * 2;
        int configMaxPoolSize = maxConcurrentSaves > 0 ? maxConcurrentSaves : defaultMaxPoolSize;
        // 确保maxPoolSize总是大于等于corePoolSize
        int maxPoolSize = Math.max(corePoolSize, configMaxPoolSize);
        int queueSize = maxQueueSize > 0 ? maxQueueSize : 10000;
        
        // 确保executorPrefix不为null
        String prefix = executorPrefix != null ? executorPrefix : "excel-task-";
        
        this.saveThreadPool = new ThreadPoolExecutor(
            corePoolSize,
            maxPoolSize,
            60L,
            TimeUnit.SECONDS,
            new LinkedBlockingQueue<>(queueSize),
            new ThreadFactory() {
                private final AtomicInteger threadNumber = new AtomicInteger(1);
                @Override
                public Thread newThread(Runnable r) {
                    Thread thread = new Thread(r, prefix + threadNumber.getAndIncrement());
                    thread.setDaemon(true);
                    return thread;
                }
            },
            new ThreadPoolExecutor.CallerRunsPolicy()
        );
    }

    @Override
    @Transactional(rollbackFor = Exception.class)
    public int batchSave(List<User> users) throws ExecutionException, InterruptedException {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        
        // 记录接收到的数据量
        log.info("开始批量保存用户数据，总条数：{}", users.size());
        
        // 验证数据是否有效
        if (users == null || users.isEmpty()) {
            log.warn("批量保存用户数据失败：数据列表为空");
            return 0;
        }
        
        // 检查第一条数据是否有效
        User firstUser = users.get(0);
        log.debug("第一条用户数据：username={}, name={}, email={}", 
                firstUser.getUsername(), firstUser.getName(), firstUser.getEmail());
        
        // 设置创建时间和更新时间
        LocalDateTime now = LocalDateTime.now();
        users.forEach(user -> {
            user.setCreateTime(now);
            user.setUpdateTime(now);
        });
        
        int totalSaved = 0;
        
        try {
            if (concurrentSaveEnabled && users.size() > PARALLEL_BATCH_SIZE) {
                // 使用并行保存
                totalSaved = parallelBatchSave(users);
            } else {
                // 使用传统的串行保存
                totalSaved = serialBatchSave(users);
            }
            
            stopWatch.stop();
            log.info("批量保存完成，共保存{}条数据，耗时{}ms，并行模式：{}", 
                    totalSaved, stopWatch.getTotalTimeMillis(), concurrentSaveEnabled && users.size() > PARALLEL_BATCH_SIZE);
        } catch (Exception e) {
            log.error("批量保存用户数据异常：", e);
            // 异常会触发事务回滚
            throw e;
        }
        
        return totalSaved;
    }
    
    /**
     * 并行批量保存数据
     */
    private int parallelBatchSave(List<User> users) throws InterruptedException, ExecutionException {
        int totalSaved = 0;
        
        // 计算并行处理的批次数
        int batchCount = (int) Math.ceil((double) users.size() / PARALLEL_BATCH_SIZE);
        log.info("并行处理批次数：{}", batchCount);
        
        // 使用CompletableFuture实现并行保存
        List<CompletableFuture<Integer>> futures = new ArrayList<>();
        
        for (int i = 0; i < batchCount; i++) {
            final int batchIndex = i;
            int start = i * PARALLEL_BATCH_SIZE;
            int end = Math.min(start + PARALLEL_BATCH_SIZE, users.size());
            List<User> batchList = new ArrayList<>(users.subList(start, end)); // 创建副本以避免并发修改问题
            
            // 提交并行任务
            CompletableFuture<Integer> future = CompletableFuture.supplyAsync(() -> {
                try {
                    // 只记录重要的批次日志，减少日志开销
                    if (batchIndex % 2 == 0) {
                        log.info("开始并行处理批次{}，数据量：{}", batchIndex + 1, batchList.size());
                    }
                    
                    // 直接调用MyBatis-Plus的批量保存，避免额外的循环开销
                    boolean saved = this.saveBatch(batchList, BATCH_SIZE);
                    
                    // 只记录重要的批次日志
                    if (batchIndex % 2 == 0) {
                        log.info("并行批次{}处理完成", batchIndex + 1);
                    }
                    
                    return saved ? batchList.size() : 0;
                } catch (Exception e) {
                    log.error("并行批次{}处理异常：", batchIndex + 1, e);
                    throw new CompletionException(e);  // 包装为CompletionException
                }
            }, saveThreadPool);
            
            futures.add(future);
        }
        
        // 等待所有并行任务完成并汇总结果
        CompletableFuture<Void> allOf = CompletableFuture.allOf(
                futures.toArray(new CompletableFuture[0])
        );
        
        // 收集所有结果
        totalSaved = allOf.thenApply(v -> 
            futures.stream()
                .map(CompletableFuture::join)
                .mapToInt(Integer::intValue)
                .sum()
        ).join();
        
        return totalSaved;
    }
    
    /**
     * 传统的串行批量保存数据
     */
    private int serialBatchSave(List<User> users) {
        int totalSaved = 0;
        int batchCount = (int) Math.ceil((double) users.size() / BATCH_SIZE);
        
        for (int i = 0; i < batchCount; i++) {
            int start = i * BATCH_SIZE;
            int end = Math.min(start + BATCH_SIZE, users.size());
            List<User> batchList = users.subList(start, end);
            
            boolean saved = this.saveBatch(batchList, BATCH_SIZE);
            log.info("串行批次{}保存结果：{}，保存数量：{}", i+1, saved, batchList.size());
            
            if (saved) {
                totalSaved += batchList.size();
            } else {
                log.warn("串行批次{}保存失败，但未抛出异常", i+1);
            }
        }
        
        return totalSaved;
    }

    @Override
    @Async
    public Future<Integer> asyncBatchSave(List<User> users) {
        try {
            int result = this.batchSave(users);
            return new AsyncResult<>(result);
        } catch (ExecutionException | InterruptedException e) {
            log.error("异步批量保存用户数据异常：", e);
            throw new CompletionException(e); // 包装为运行时异常
        }
    }

    @Override
    public List<User> selectPage(int pageNum, int pageSize) {
        long offset = (long) (pageNum - 1) * pageSize;
        return userMapper.selectList(null)
                .stream()
                .skip(offset)
                .limit(pageSize)
                .collect(Collectors.toList());
    }

    @Override
    @Transactional(rollbackFor = Exception.class)
    public void generateTestData(int count) {
        long startTime = System.currentTimeMillis();
        log.info("开始生成测试数据，总条数：{}", count);
        
        List<User> users = new ArrayList<>();
        for (int i = 0; i < count; i++) {
            User user = new User();
            user.setUsername("user_" + i);
            user.setName("测试用户" + i);
            user.setEmail("user_" + i + "@example.com");
            user.setPhone("138" + String.format("%08d", i));
            user.setAge(random.nextInt(50) + 18);
            user.setCreateTime(LocalDateTime.now());
            user.setUpdateTime(LocalDateTime.now());
            users.add(user);
            
            // 每达到批次大小就保存一次
            if (users.size() >= BATCH_SIZE) {
                this.saveBatch(users, BATCH_SIZE);
                users.clear();
            }
        }
        
        // 保存剩余的数据
        if (!users.isEmpty()) {
            this.saveBatch(users, BATCH_SIZE);
        }
        
        long endTime = System.currentTimeMillis();
        log.info("测试数据生成完成，耗时{}ms", (endTime - startTime));
    }
}