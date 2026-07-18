package com.lframework.starter.web.core.utils;

import lombok.Data;

import java.util.List;

@Data
public class MultiSheetsData<T> {
    private List<T> data;
    private Class<T> headClazz;
    private String sheetName;
}
