package com.paperformat.server.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.scheduling.concurrent.ThreadPoolTaskExecutor;

import java.util.concurrent.Executor;

/**
 * 后台格式处理线程池配置。
 */
@Configuration
public class TaskExecutorConfig {
    /**
     * 限制并发处理数量，避免多个大 DOCX 同时处理时占满服务器资源。
     */
    @Bean(name = "formatTaskExecutor")
    public Executor formatTaskExecutor() {
        ThreadPoolTaskExecutor executor = new ThreadPoolTaskExecutor();
        executor.setCorePoolSize(2);
        executor.setMaxPoolSize(4);
        executor.setQueueCapacity(20);
        executor.setThreadNamePrefix("format-task-");
        executor.initialize();
        return executor;
    }
}
