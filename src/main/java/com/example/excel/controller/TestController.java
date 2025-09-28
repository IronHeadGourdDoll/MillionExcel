package com.example.excel.controller;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * Simple test controller to diagnose application startup issues
 */
@RestController
@RequestMapping("/api/test")
public class TestController {
    
    @GetMapping("/ping")
    public String ping() {
        return "Application is running!";
    }
}