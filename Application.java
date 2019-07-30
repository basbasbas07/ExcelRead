/*
 * Develoed by Barış Bülbül on 23.07.2019 14:57.
 * Name of the current project excelread.
 * Name of the current module ExcelRead.
 * Last modified 23.07.2019 11:13.
 * StjBarisB.
 * Copyright (c) 2019. All rights reserved.
 */

package com.gt.shift;

        import org.springframework.boot.SpringApplication;
        import org.springframework.boot.SpringBootConfiguration;

@SpringBootConfiguration
public class Application {

    public static void main(String[] args) {
        System.out.println("APP STARTED!");
        SpringApplication.run(Application.class, args);
    }
}
