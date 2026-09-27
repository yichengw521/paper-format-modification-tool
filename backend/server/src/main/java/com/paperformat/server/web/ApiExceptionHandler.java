package com.paperformat.server.web;

import org.springframework.http.HttpStatus;
import org.springframework.http.ProblemDetail;
import org.springframework.web.bind.MissingServletRequestParameterException;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.multipart.MaxUploadSizeExceededException;

import java.io.IOException;
import java.util.NoSuchElementException;

/**
 * 统一把服务层异常转换成 ProblemDetail JSON 响应。
 * 这是一个全局异常处理器，通过@RestControllerAdvice注解实现。
 * 它捕获各种异常并将它们转换为符合RFC 7807标准的ProblemDetail格式响应。
 */
@RestControllerAdvice
public class ApiExceptionHandler {
    /**
     * 处理非法参数异常和缺少请求参数异常
     * @param exception 捕获的异常对象
     * @return 返回状态码为400的ProblemDetail
     */
    @ExceptionHandler({IllegalArgumentException.class, MissingServletRequestParameterException.class})
    // 处理非法参数和缺少请求参数的方法，返回400错误响应
    ProblemDetail badRequest(Exception exception) {
        return problem(HttpStatus.BAD_REQUEST, exception.getMessage());
    }

    /**
     * 处理资源未找到异常
     * @param exception 捕获的NoSuchElementException异常
     * @return 返回状态码为404的ProblemDetail
     */
    @ExceptionHandler(NoSuchElementException.class)
    // 处理资源未找到异常的方法，返回404错误响应
    ProblemDetail notFound(NoSuchElementException exception) {
        return problem(HttpStatus.NOT_FOUND, exception.getMessage());
    }

    /**
     * 处理冲突状态异常
     * @param exception 捕获的IllegalStateException异常
     * @return 返回状态码为409的ProblemDetail
     */
    @ExceptionHandler(IllegalStateException.class)
    // 处理冲突状态异常的方法，返回409错误响应
    ProblemDetail conflict(IllegalStateException exception) {
        return problem(HttpStatus.CONFLICT, exception.getMessage());
    }

    /**
     * 处理文件上传大小超限异常
     * @param exception 捕获的MaxUploadSizeExceededException异常
     * @return 返回状态码为413的ProblemDetail，提示上传文件超出配置大小限制
     */
    @ExceptionHandler(MaxUploadSizeExceededException.class)
    // 处理文件上传大小超限异常的方法，返回413错误响应
    ProblemDetail tooLarge(MaxUploadSizeExceededException exception) {
        return problem(HttpStatus.PAYLOAD_TOO_LARGE, "Uploaded files exceed the configured size limit.");
    }

    /**
     * 处理文件IO异常
     * @param exception 捕获的IOException异常
     * @return 返回状态码为500的ProblemDetail，提示文件处理失败
     */
    @ExceptionHandler(IOException.class)
    // 处理文件IO异常的方法，返回500错误响应
    ProblemDetail ioError(IOException exception) {
        return problem(HttpStatus.INTERNAL_SERVER_ERROR, "File processing failed.");
    }

    /**
     * Spring 6 的 ProblemDetail 符合 RFC 7807，前端统一读取 detail 字段即可。
     */
    // 创建ProblemDetail对象的私有辅助方法
    private static ProblemDetail problem(HttpStatus status, String detail) {
        // 根据HTTP状态码和错误详情创建ProblemDetail对象
        ProblemDetail problem = ProblemDetail.forStatusAndDetail(status, detail == null ? status.getReasonPhrase() : detail);
        // 设置ProblemDetail的标题为状态码对应的描述
        problem.setTitle(status.getReasonPhrase());
        return problem;
    }
}
