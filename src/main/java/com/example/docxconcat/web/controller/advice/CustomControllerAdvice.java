package com.example.docxconcat.web.controller.advice;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.ResponseStatus;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.context.request.WebRequest;

import java.util.Date;

/**
 * @author ogbozoyan
 * @date 08.07.2023
 */
@RestControllerAdvice
public class CustomControllerAdvice {
    private static final Integer INTERNAL_SERVER_ERROR = 500; //INTERNAL_SERVER_ERROR
    private static final Integer NOT_FOUND = 404;

    private static final Integer FORBIDDEN = 403;
    private static final Integer BAD_REQUEST = 400;

    @ExceptionHandler(RuntimeException.class)
    @ResponseStatus(HttpStatus.INTERNAL_SERVER_ERROR)
    public ResponseEntity<CustomErrorMessage> runtimeException(RuntimeException er, WebRequest req) {
        CustomErrorMessage errorMessage = new CustomErrorMessage(
                INTERNAL_SERVER_ERROR,
                new Date(),
                er.getMessage(),
                req.getDescription(false)
        );
        return new ResponseEntity<>(errorMessage, HttpStatus.INTERNAL_SERVER_ERROR);
    }
}
