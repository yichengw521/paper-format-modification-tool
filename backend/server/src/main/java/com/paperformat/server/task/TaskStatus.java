package com.paperformat.server.task;

/**
 * 任务生命周期状态。
 */
public enum TaskStatus {
    ANALYZING,
    AWAITING_CONFIRMATION,
    QUEUED,
    PROCESSING,
    COMPLETED,
    FAILED
}
