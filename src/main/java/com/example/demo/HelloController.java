package com.example.demo;

import org.springframework.stereotype.Controller;

@Controller
public class HelloController {
    public String hello() {
        return "index";
    }
}