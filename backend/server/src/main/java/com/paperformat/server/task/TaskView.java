package com.paperformat.server.task;

import java.time.OffsetDateTime;
import java.util.Map;

/**
 * 返回给前端的任务状态视图。
 *
 * <p>这里不暴露服务器本地路径，只暴露可访问的 API 链接和摘要信息。</p>
 */
public record TaskView(
        String id,
        TaskStatus status,
        OffsetDateTime createdAt,
        OffsetDateTime updatedAt,
        String templateFileName,
        String documentFileName,
        String resultFileName,
        String reportFileName,
        String analysisFileName,
        Integer modifications,
        Integer tables,
        Integer sections,
        String error,
        Map<String, String> links
) {
}
