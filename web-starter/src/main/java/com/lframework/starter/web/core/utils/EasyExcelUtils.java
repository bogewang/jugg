package com.lframework.starter.web.core.utils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.handler.WriteHandler;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.column.AbstractColumnWidthStyleStrategy;
import com.alibaba.excel.write.style.column.SimpleColumnWidthStyleStrategy;
import com.alibaba.excel.write.style.row.SimpleRowHeightStyleStrategy;
import com.google.common.collect.Lists;

import com.lframework.starter.common.utils.DateUtil;
import com.lframework.starter.common.utils.ReflectUtil;
import com.lframework.starter.web.core.annotations.excel.ExcelRequired;
import com.lframework.starter.web.core.components.excel.ExcelHorizontalCellStyleStrategy;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;

@Slf4j
public class EasyExcelUtils {
    // 默认列宽
    public static final int DEFAULT_COLUMN_WIDTH = 15;
    // 默认表头高度
    public static final short DEFAULT_HEADER_HEIGHT = 30;
    // 默认内容高度
    public static final short DEFAULT_CONTENT_HEIGHT = 20;

    private static final String DEFAULT_SHEET_NAME = "Sheet1";
    /**
     * 默认列宽策略
     */
    private static final WriteHandler DEFAULT_COLUMN_WIDTH_STYLE_STRATEGY = new SimpleColumnWidthStyleStrategy(DEFAULT_COLUMN_WIDTH);

    /**
     * 同步无模型读（默认读取sheet0,从第2行开始读）
     *
     * @param filePath excel文件的绝对路径
     */
    public static List<Map<Integer, String>> syncRead(String filePath) {
        return EasyExcelFactory.read(filePath).sheet().doReadSync();
    }

    /**
     * 同步无模型读（默认读取sheet0,从第2行开始读）
     *
     * @param inputStream excel文件的输入流
     */
    public static List<Map<Integer, String>> syncRead(InputStream inputStream) {
        return EasyExcelFactory.read(inputStream).sheet().doReadSync();
    }

    /**
     * 同步无模型读（默认读取sheet0,从第2行开始读）
     *
     * @param file excel文件
     */
    public static List<Map<Integer, String>> syncRead(File file) {
        return EasyExcelFactory.read(file).sheet().doReadSync();
    }

    /**
     * 同步无模型读（自定义读取sheetX，从第2行开始读）
     *
     * @param filePath excel文件的绝对路径
     * @param sheetNo  sheet页号，从0开始
     */
    public static List<Map<Integer, String>> syncRead(String filePath, Integer sheetNo) {
        return EasyExcelFactory.read(filePath).sheet(sheetNo).doReadSync();
    }

    /**
     * 同步无模型读（自定义读取sheetX，从第2行开始读）
     *
     * @param inputStream excel文件的输入流
     * @param sheetNo     sheet页号，从0开始
     */
    public static List<Map<Integer, String>> syncRead(InputStream inputStream, Integer sheetNo) {
        return EasyExcelFactory.read(inputStream).sheet(sheetNo).doReadSync();
    }

    /**
     * 同步无模型读（自定义读取sheetX，从第2行开始读）
     *
     * @param file    excel文件
     * @param sheetNo sheet页号，从0开始
     */
    public static List<Map<Integer, String>> syncRead(File file, Integer sheetNo) {
        return EasyExcelFactory.read(file).sheet(sheetNo).doReadSync();
    }

    /**
     * 同步无模型读（指定sheet和表头占的行数）
     *
     * @param filePath
     * @param sheetNo    sheet页号，从0开始
     * @param headRowNum 表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static List<Map<Integer, String>> syncRead(String filePath, Integer sheetNo, Integer headRowNum) {
        return EasyExcelFactory.read(filePath).sheet(sheetNo).headRowNumber(headRowNum).doReadSync();
    }

    /**
     * 同步无模型读（指定sheet和表头占的行数）
     *
     * @param inputStream
     * @param sheetNo     sheet页号，从0开始
     * @param headRowNum  表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static List<Map<Integer, String>> syncRead(InputStream inputStream, Integer sheetNo, Integer headRowNum) {
        return EasyExcelFactory.read(inputStream).sheet(sheetNo).headRowNumber(headRowNum).doReadSync();
    }

    /**
     * 同步无模型读（指定sheet和表头占的行数）
     *
     * @param file
     * @param sheetNo    sheet页号，从0开始
     * @param headRowNum 表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static List<Map<Integer, String>> syncRead(File file, Integer sheetNo, Integer headRowNum) {
        return EasyExcelFactory.read(file).sheet(sheetNo).headRowNumber(headRowNum).doReadSync();
    }

    /**
     * 同步按模型读（默认读取sheet0,从第2行开始读）
     *
     * @param filePath
     * @param clazz    模型的类类型（excel数据会按该类型转换成对象）
     */
    public static <T> List<T> syncReadModel(String filePath, Class<T> clazz) {
        return EasyExcelFactory.read(filePath).sheet().head(clazz).doReadSync();
    }

    /**
     * 同步按模型读（默认读取sheet0,从第2行开始读）
     *
     * @param inputStream
     * @param clazz       模型的类类型（excel数据会按该类型转换成对象）
     */
    public static <T> List<T> syncReadModel(InputStream inputStream, Class<T> clazz) {
        return EasyExcelFactory.read(inputStream).sheet().head(clazz).doReadSync();
    }

    /**
     * 同步按模型读（默认读取sheet0,从第2行开始读）
     *
     * @param file
     * @param clazz 模型的类类型（excel数据会按该类型转换成对象）
     */
    public static <T> List<T> syncReadModel(File file, Class<T> clazz) {
        return EasyExcelFactory.read(file).sheet().head(clazz).doReadSync();
    }

    /**
     * 同步按模型读（默认表头占一行，从第2行开始读）
     *
     * @param filePath
     * @param clazz    模型的类类型（excel数据会按该类型转换成对象）
     * @param sheetNo  sheet页号，从0开始
     */
    public static <T> List<T> syncReadModel(String filePath, Class<T> clazz, Integer sheetNo) {
        return EasyExcelFactory.read(filePath).sheet(sheetNo).head(clazz).doReadSync();
    }

    /**
     * 同步按模型读（默认表头占一行，从第2行开始读）
     *
     * @param inputStream
     * @param clazz       模型的类类型（excel数据会按该类型转换成对象）
     * @param sheetNo     sheet页号，从0开始
     */
    public static <T> List<T> syncReadModel(InputStream inputStream, Class<T> clazz, Integer sheetNo) {
        return EasyExcelFactory.read(inputStream).sheet(sheetNo).head(clazz).doReadSync();
    }

    /**
     * 同步按模型读（默认表头占一行，从第2行开始读）
     *
     * @param file
     * @param clazz   模型的类类型（excel数据会按该类型转换成对象）
     * @param sheetNo sheet页号，从0开始
     */
    public static <T> List<T> syncReadModel(File file, Class<T> clazz, Integer sheetNo) {
        return EasyExcelFactory.read(file).sheet(sheetNo).head(clazz).doReadSync();
    }

    /**
     * 同步按模型读（指定sheet和表头占的行数）
     *
     * @param filePath
     * @param clazz      模型的类类型（excel数据会按该类型转换成对象）
     * @param sheetNo    sheet页号，从0开始
     * @param headRowNum 表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static <T> List<T> syncReadModel(String filePath, Class<T> clazz, Integer sheetNo, Integer headRowNum) {
        return EasyExcelFactory.read(filePath).sheet(sheetNo).headRowNumber(headRowNum).head(clazz).doReadSync();
    }

    /**
     * 同步按模型读（指定sheet和表头占的行数）
     *
     * @param inputStream
     * @param clazz       模型的类类型（excel数据会按该类型转换成对象）
     * @param sheetNo     sheet页号，从0开始
     * @param headRowNum  表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static <T> List<T> syncReadModel(InputStream inputStream, Class<T> clazz, Integer sheetNo, Integer headRowNum) {
        return EasyExcelFactory.read(inputStream).sheet(sheetNo).headRowNumber(headRowNum).head(clazz).doReadSync();
    }

    /**
     * 同步按模型读（指定sheet和表头占的行数）
     *
     * @param file
     * @param clazz      模型的类类型（excel数据会按该类型转换成对象）
     * @param sheetNo    sheet页号，从0开始
     * @param headRowNum 表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static <T> List<T> syncReadModel(File file, Class<T> clazz, Integer sheetNo, Integer headRowNum) {
        return EasyExcelFactory.read(file).sheet(sheetNo).headRowNumber(headRowNum).head(clazz).doReadSync();
    }

    /**
     * 异步无模型读（默认读取sheet0,从第2行开始读）
     *
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param filePath      表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static <T> void asyncRead(String filePath, AnalysisEventListener<T> excelListener) {
        EasyExcelFactory.read(filePath, excelListener).sheet().doRead();
    }

    /**
     * 异步无模型读（默认读取sheet0,从第2行开始读）
     *
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param inputStream   表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static <T> void asyncRead(InputStream inputStream, AnalysisEventListener<T> excelListener) {
        EasyExcelFactory.read(inputStream, excelListener).sheet().doRead();
    }

    /**
     * 异步无模型读（默认读取sheet0,从第2行开始读）
     *
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param file          表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static <T> void asyncRead(File file, AnalysisEventListener<T> excelListener) {
        EasyExcelFactory.read(file, excelListener).sheet().doRead();
    }

    /**
     * 异步无模型读（默认表头占一行，从第2行开始读）
     *
     * @param filePath      表头占的行数，从0开始（如果要连表头一起读出来则传0）
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param sheetNo       sheet页号，从0开始
     */
    public static <T> void asyncRead(String filePath, AnalysisEventListener<T> excelListener, Integer sheetNo) {
        EasyExcelFactory.read(filePath, excelListener).sheet(sheetNo).doRead();
    }

    /**
     * 异步无模型读（默认表头占一行，从第2行开始读）
     *
     * @param inputStream   表头占的行数，从0开始（如果要连表头一起读出来则传0）
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param sheetNo       sheet页号，从0开始
     */
    public static <T> void asyncRead(InputStream inputStream, AnalysisEventListener<T> excelListener, Integer sheetNo) {
        EasyExcelFactory.read(inputStream, excelListener).sheet(sheetNo).doRead();
    }

    /**
     * 异步无模型读（默认表头占一行，从第2行开始读）
     *
     * @param file          表头占的行数，从0开始（如果要连表头一起读出来则传0）
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param sheetNo       sheet页号，从0开始
     */
    public static <T> void asyncRead(File file, AnalysisEventListener<T> excelListener, Integer sheetNo) {
        EasyExcelFactory.read(file, excelListener).sheet(sheetNo).doRead();
    }

    /**
     * 异步无模型读（指定sheet和表头占的行数）
     *
     * @param filePath
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param sheetNo       sheet页号，从0开始
     * @param headRowNum    表头占的行数，从0开始（如果要连表头一起读出来则传0）
     * @return
     */
    public static <T> void asyncRead(String filePath, AnalysisEventListener<T> excelListener, Integer sheetNo,
                                     Integer headRowNum) {
        EasyExcelFactory.read(filePath, excelListener).sheet(sheetNo).headRowNumber(headRowNum).doRead();
    }

    /**
     * 异步无模型读（指定sheet和表头占的行数）
     *
     * @param inputStream
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param sheetNo       sheet页号，从0开始
     * @param headRowNum    表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static <T> void asyncRead(InputStream inputStream, AnalysisEventListener<T> excelListener, Integer sheetNo,
                                     Integer headRowNum) {
        EasyExcelFactory.read(inputStream, excelListener).sheet(sheetNo).headRowNumber(headRowNum).doRead();
    }

    /**
     * 异步无模型读（指定sheet和表头占的行数）
     *
     * @param file
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param sheetNo       sheet页号，从0开始
     * @param headRowNum    表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static <T> void asyncRead(File file, AnalysisEventListener<T> excelListener, Integer sheetNo, Integer headRowNum) {
        EasyExcelFactory.read(file, excelListener).sheet(sheetNo).headRowNumber(headRowNum).doRead();
    }

    /**
     * 异步按模型读取（默认读取sheet0,从第2行开始读）
     *
     * @param filePath
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param clazz         模型的类类型（excel数据会按该类型转换成对象）
     */
    public static <T> void asyncReadModel(String filePath, AnalysisEventListener<T> excelListener, Class<T> clazz) {
        EasyExcelFactory.read(filePath, clazz, excelListener).sheet().doRead();
    }

    /**
     * 异步按模型读取（默认读取sheet0,从第2行开始读）
     *
     * @param inputStream
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param clazz         模型的类类型（excel数据会按该类型转换成对象）
     */
    public static <T> void asyncReadModel(InputStream inputStream, AnalysisEventListener<T> excelListener,
                                          Class<T> clazz) {
        EasyExcelFactory.read(inputStream, clazz, excelListener).sheet().doRead();
    }

    /**
     * 异步按模型读取（默认读取sheet0,从第2行开始读）
     *
     * @param file
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param clazz         模型的类类型（excel数据会按该类型转换成对象）
     */
    public static <T> void asyncReadModel(File file, AnalysisEventListener<T> excelListener, Class<T> clazz) {
        EasyExcelFactory.read(file, clazz, excelListener).sheet().doRead();
    }

    /**
     * 异步按模型读取（默认表头占一行，从第2行开始读）
     *
     * @param filePath
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param clazz         模型的类类型（excel数据会按该类型转换成对象）
     * @param sheetNo       sheet页号，从0开始
     */
    public static <T> void asyncReadModel(String filePath, AnalysisEventListener<T> excelListener, Class<T> clazz,
                                          Integer sheetNo) {
        EasyExcelFactory.read(filePath, clazz, excelListener).sheet(sheetNo).doRead();
    }

    /**
     * 异步按模型读取（默认表头占一行，从第2行开始读）
     *
     * @param inputStream
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param clazz         模型的类类型（excel数据会按该类型转换成对象）
     * @param sheetNo       sheet页号，从0开始
     */
    public static <T> void asyncReadModel(InputStream inputStream, AnalysisEventListener<T> excelListener, Class<T> clazz,
                                          Integer sheetNo) {
        EasyExcelFactory.read(inputStream, clazz, excelListener).sheet(sheetNo).doRead();
    }

    /**
     * 异步按模型读取（默认表头占一行，从第2行开始读）
     *
     * @param file
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param clazz         模型的类类型（excel数据会按该类型转换成对象）
     * @param sheetNo       sheet页号，从0开始
     */
    public static <T> void asyncReadModel(File file, AnalysisEventListener<T> excelListener, Class<T> clazz,
                                          Integer sheetNo) {
        EasyExcelFactory.read(file, clazz, excelListener).sheet(sheetNo).doRead();
    }

    /**
     * 异步按模型读取
     *
     * @param filePath
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param clazz         模型的类类型（excel数据会按该类型转换成对象）
     * @param sheetNo       sheet页号，从0开始
     * @param headRowNum    表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static <T> void asyncReadModel(String filePath, AnalysisEventListener<T> excelListener, Class<T> clazz,
                                          Integer sheetNo, Integer headRowNum) {
        EasyExcelFactory.read(filePath, clazz, excelListener).sheet(sheetNo).headRowNumber(headRowNum).doRead();
    }

    /**
     * 异步按模型读取
     *
     * @param inputStream
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param clazz         模型的类类型（excel数据会按该类型转换成对象）
     * @param sheetNo       sheet页号，从0开始
     * @param headRowNum    表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static <T> void asyncReadModel(InputStream inputStream, AnalysisEventListener<T> excelListener, Class<T> clazz,
                                          Integer sheetNo, Integer headRowNum) {
        EasyExcelFactory.read(inputStream, clazz, excelListener).sheet(sheetNo).headRowNumber(headRowNum).doRead();
    }

    /**
     * 异步按模型读取
     *
     * @param file
     * @param excelListener 监听器，在监听器中可以处理行数据LinkedHashMap，表头数据，异常处理等
     * @param clazz         模型的类类型（excel数据会按该类型转换成对象）
     * @param sheetNo       sheet页号，从0开始
     * @param headRowNum    表头占的行数，从0开始（如果要连表头一起读出来则传0）
     */
    public static <T> void asyncReadModel(File file, AnalysisEventListener<T> excelListener, Class<T> clazz,
                                          Integer sheetNo, Integer headRowNum) {
        EasyExcelFactory.read(file, clazz, excelListener).sheet(sheetNo).headRowNumber(headRowNum).doRead();
    }

    /**
     * 无模板写文件
     *
     * @param filePath
     * @param head     表头数据
     * @param data     表内容数据
     */
    public static void write(String filePath, List<List<String>> head, List<List<Object>> data) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(filePath).head(head).sheet();

        customCellStyle(builder, null, null, data);
    }

    /**
     * 无模板写文件
     *
     * @param outputStream
     * @param head         表头数据
     * @param data         表内容数据
     */
    public static void write(OutputStream outputStream, List<List<String>> head, List<List<Object>> data) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(outputStream).head(head).sheet();
        customCellStyle(builder, null, null, data);
    }

    /**
     * 无模板写文件
     *
     * @param filePath
     * @param head     表头数据
     * @param data     表内容数据
     * @param sheetNo  sheet页号，从0开始
     */
    public static void write(String filePath, List<List<String>> head, List<List<Object>> data, Integer sheetNo) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(filePath).head(head).sheet(sheetNo);

        customCellStyle(builder, null, null, data);
    }

    /**
     * 无模板写文件
     *
     * @param outputStream
     * @param head         表头数据
     * @param data         表内容数据
     * @param sheetNo      sheet页号，从0开始
     */
    public static void write(OutputStream outputStream, List<List<String>> head, List<List<Object>> data, Integer sheetNo) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(outputStream).head(head).sheet(sheetNo);

        customCellStyle(builder, null, null, data);
    }

    /**
     * 无模板写文件
     *
     * @param filePath
     * @param head      表头数据
     * @param data      表内容数据
     * @param sheetNo   sheet页号，从0开始
     * @param sheetName sheet名称
     */
    public static void write(String filePath, List<List<String>> head, List<List<Object>> data, Integer sheetNo, String sheetName) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(filePath).head(head).sheet(sheetNo, sheetName);
        customCellStyle(builder, null, null, data);
    }

    /**
     * 无模板写文件
     *
     * @param outputStream
     * @param head         表头数据
     * @param data         表内容数据
     * @param sheetNo      sheet页号，从0开始
     * @param sheetName    sheet名称
     */
    public static void write(OutputStream outputStream, List<List<String>> head, List<List<Object>> data,
                             Integer sheetNo, String sheetName) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(outputStream).head(head).sheet(sheetNo, sheetName);
        customCellStyle(builder, null, null, data);
    }

    /**
     * 根据excel模板文件写入文件
     *
     * @param filePath
     * @param templateFileName
     * @param data
     */
    public static <T> void writeTemplate(String filePath, String templateFileName, List<T> data) {
        EasyExcel.write(filePath).withTemplate(templateFileName).sheet().doFill(data);
    }

    /**
     * 根据excel模板文件写入文件
     *
     * @param outputStream
     * @param templateFileName
     * @param data
     */
    public static <T> void writeTemplate(OutputStream outputStream, String templateFileName, List<T> data) {
        EasyExcel.write(outputStream).withTemplate(templateFileName).sheet().doFill(data);
    }

    /**
     * 根据excel模板文件写入文件
     *
     * @param file
     * @param templateFileName
     * @param data
     */
    public static <T> void writeTemplate(File file, String templateFileName, List<T> data) {
        EasyExcel.write(file).withTemplate(templateFileName).sheet().doFill(data);
    }

    /**
     * 根据excel模板文件写入文件
     *
     * @param filePath
     * @param templateFileName
     * @param headClazz
     * @param data
     */
    public static <T> void writeTemplate(String filePath, String templateFileName,
                                                            Class<T> headClazz, List<T> data) {
        EasyExcel.write(filePath, headClazz).withTemplate(templateFileName).sheet().doFill(data);
    }

    /**
     * 根据excel模板文件写入文件
     *
     * @param outputStream
     * @param templateFileName
     * @param headClazz
     * @param data
     */
    public static <T> void writeTemplate(OutputStream outputStream, String templateFileName,
                                                            Class<T> headClazz, List<T> data) {
        EasyExcel.write(outputStream, headClazz).withTemplate(templateFileName).sheet().doFill(data);
    }

    /**
     * 根据excel模板文件写入文件
     *
     * @param file
     * @param templateFileName
     * @param headClazz
     * @param data
     */
    public static <T> void writeTemplate(File file, String templateFileName,
                                                            Class<T> headClazz, List<T> data) {
        EasyExcel.write(file, headClazz).withTemplate(templateFileName).sheet().doFill(data);
    }

    /**
     * 按模板写文件
     *
     * @param filePath
     * @param headClazz 表头模板
     * @param data      数据
     */
    public static <T> void write(String filePath, Class<T> headClazz, List<T> data) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(filePath, headClazz).sheet();

        customCellStyle(builder, null, headClazz, data);
    }

    /**
     * 按模板写文件
     *
     * @param outputStream
     * @param headClazz    表头模板
     * @param data         数据
     */
    public static <T> void write(OutputStream outputStream, Class<T> headClazz, List<T> data) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(outputStream, headClazz).sheet();
        customCellStyle(builder, null, headClazz, data);
    }


    /**
     * 按模板写文件
     *
     * @param file
     * @param headClazz 表头模板
     * @param data      数据
     */
    public static <T> void write(File file, Class<T> headClazz, List<T> data) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(file, headClazz).sheet();
        customCellStyle(builder, null, headClazz, data);
    }

    /**
     * 按模板写文件
     *
     * @param filePath
     * @param headClazz 表头模板
     * @param data      数据
     * @param sheetNo   sheet页号，从0开始
     */
    public static <T> void write(String filePath, Class<T> headClazz, List<T> data, Integer sheetNo) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(filePath, headClazz).sheet(sheetNo);

        customCellStyle(builder, null, headClazz, data);
    }

    public static <T> void customCellStyle(ExcelWriterSheetBuilder builder,
                                                              List<WriteHandler> writeHandlers,
                                                              Class<T> headClazz, List data) {
        writeHandlers = getWriteHandlers(writeHandlers, headClazz);
        writeHandlers.forEach(builder::registerWriteHandler);

        builder.doWrite(data);
    }

    /**
     * 不执行doWrite
     * @param builder
     * @param writeHandlers
     * @param headClazz
     * @param data
     * @param <T>
     */
    public static <T> void customCellStyleNoWrite(ExcelWriterSheetBuilder builder,
                                           List<WriteHandler> writeHandlers,
                                           Class<T> headClazz, List data) {
        writeHandlers = getWriteHandlers(writeHandlers, headClazz);
        writeHandlers.forEach(builder::registerWriteHandler);

    }

    /**
     * 按模板写文件
     *
     * @param response
     * @param headClazz 表头模板
     * @param data      数据
     * @param sheetNo   sheet页号，从0开始
     */
    public static <T> void write(HttpServletResponse response, String fileName,
                                                    Class<T> headClazz, List<T> data, Integer sheetNo)
            throws IOException {
        OutputStream outputStream = formatResponse(fileName, response);
        ExcelWriterSheetBuilder builder = EasyExcel.write(outputStream, headClazz).sheet(sheetNo);

        customCellStyle(builder, null, headClazz, data);
    }

    /**
     * 按模板写文件
     *
     * @param file
     * @param headClazz 表头模板
     * @param data      数据
     * @param sheetNo   sheet页号，从0开始
     */
    public static <T> void write(File file, Class<T> headClazz, List<T> data, Integer sheetNo) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(file, headClazz).sheet(sheetNo);

        customCellStyle(builder, null, headClazz, data);
    }

    /**
     * 按模板写文件
     *
     * @param filePath
     * @param headClazz 表头模板
     * @param data      数据
     * @param sheetNo   sheet页号，从0开始
     * @param sheetName sheet名称
     */
    public static <T> void write(String filePath, Class<T> headClazz, List<T> data,
                                                    Integer sheetNo, String sheetName) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(filePath, headClazz).sheet(sheetNo, sheetName);

        customCellStyle(builder, null, headClazz, data);
    }

    /**
     * 按模板写文件
     *
     * @param outputStream
     * @param headClazz    表头模板
     * @param data         数据
     * @param sheetNo      sheet页号，从0开始
     * @param sheetName    sheet名称
     */
    public static <T> void write(OutputStream outputStream, Class<T> headClazz, List<T> data,
                                                    Integer sheetNo, String sheetName) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(outputStream, headClazz).sheet(sheetNo, sheetName);

        customCellStyle(builder, null, headClazz, data);
    }

    /**
     * 按模板写文件
     *
     * @param file
     * @param headClazz 表头模板
     * @param data      数据
     * @param sheetNo   sheet页号，从0开始
     * @param sheetName sheet名称
     */
    public static <T> void write(File file, Class<T> headClazz, List<T> data,
                                                    Integer sheetNo, String sheetName) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(file, headClazz).sheet(sheetNo, sheetName);

        customCellStyle(builder, null, headClazz, data);
    }

    /**
     * 按模板写文件
     *
     * @param filePath
     * @param headClazz    表头模板
     * @param data         数据
     * @param writeHandler 自定义的处理器，比如设置table样式，设置超链接、单元格下拉框等等功能都可以通过这个实现（需要注册多个则自己通过链式去调用）
     * @param sheetNo      sheet页号，从0开始
     * @param sheetName    sheet名称
     */
    public static <T> void write(String filePath, Class<T> headClazz, List<T> data,
                                                    WriteHandler writeHandler, Integer sheetNo, String sheetName) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(filePath, headClazz).sheet(sheetNo, sheetName);

        customCellStyle(builder, Lists.newArrayList(writeHandler), headClazz, data);
    }


    /**
     * 多个sheet页的数据链式写入
     * ExcelUtil.writeWithSheets(file)
     * .writeModel(ExcelModel.class, excelModelList, "sheetName1")
     * .write(headData, data,"sheetName2")
     * .finish();
     *
     * @param file
     */
    public static EasyExcelWriterFactory writeWithSheets(File file) {
        return new EasyExcelWriterFactory(file);
    }

    /**
     * 多个sheet页的数据链式写入
     * ExcelUtil.writeWithSheets(filePath)
     * .writeModel(ExcelModel.class, excelModelList, "sheetName1")
     * .write(headData, data,"sheetName2")
     * .finish();
     *
     * @param filePath
     */
    public static EasyExcelWriterFactory writeWithSheets(String filePath) {
        EasyExcelWriterFactory excelWriter = new EasyExcelWriterFactory(filePath);
        return excelWriter;
    }

    /**
     * 读取Excel文件
     * 使用EasyExcel读取Excel文件并返回读取构建器
     *
     * @param file     Excel文件，不能为null
     * @param clazz    目标类型，不能为null
     * @param listener 读取监听器，不能为null
     * @param <T>      目标类型泛型
     * @return Excel读取构建器
     */
    public static <T> ExcelReaderBuilder read(MultipartFile file, Class<T> clazz, ReadListener<T> listener) {
        try {
            return EasyExcel.read(file.getInputStream(), clazz, listener);
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException(e.getMessage());
        }
    }


    /**
     * 按模板写文件
     *
     * @param fileName
     * @param headClazz 表头模板
     * @param data      数据
     */
    public static <T> void write(HttpServletResponse response, String fileName, Class<T> headClazz, List<T> data) {
        write(response, fileName, null, headClazz, data);
    }

    /**
     * 设置Excel头
     *
     * @param headList Excel头信息
     * @return
     */
    private static List<List<String>> head(List<String> headList) {
        List<List<String>> list = new ArrayList<>();
        for (String value : headList) {
            List<String> head = new ArrayList<>();
            head.add(value);
            list.add(head);
        }
        return list;
    }

    /**
     * 动态表头写文件
     *
     * @param fileName
     * @param headerMap 表头模板()
     * @param data      数据
     */
    public static <T> void write(HttpServletResponse response, String fileName, Map<String, String> headerMap, List<T> data) {
        try {
            List<String> headers = Lists.newArrayList();
            List<String> fields = Lists.newArrayList();

            headerMap.forEach((k, v) -> {
                fields.add(k);
                headers.add(v);
            });

            List<List<?>> datas = dataList(data, fields);

            OutputStream outputStream = formatResponse(fileName, response);
            ExcelWriterSheetBuilder builder = EasyExcel.write(outputStream)
                    .excelType(ExcelTypeEnum.XLSX)
                    .head(head(headers))
                    .sheet()
                    .sheetName(DEFAULT_SHEET_NAME);

            customCellStyle(builder, null, null, datas);
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException(e);
        }
    }

    /**
     * 动态表头写文件
     *
     * @param fileName
     * @param headerMap 表头模板()
     * @param data      数据
     */
    public static <T> void writeNoModel(HttpServletResponse response, String fileName, Map<String, String> headerMap, List<Map<String, String>> data) {
        try {
            List<String> headers = Lists.newArrayList();
            List<String> fields = Lists.newArrayList();

            headerMap.forEach((k, v) -> {
                fields.add(k);
                headers.add(v);
            });

            List<List<?>> datas = dataList(data, fields);

            OutputStream outputStream = formatResponse(fileName, response);
            ExcelWriterSheetBuilder builder = EasyExcel.write(outputStream)
                    .excelType(ExcelTypeEnum.XLSX)
                    .head(head(headers))
                    .sheet()
                    .sheetName(DEFAULT_SHEET_NAME);

            customCellStyle(builder, null, null, datas);
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException(e);
        }
    }

    /**
     * 设置表格信息
     *
     * @param dataList 查询出的数据
     * @param fileList 需要显示的字段
     * @return
     */
    private static List<List<?>> dataList(List<?> dataList, List<String> fileList) {
        List<List<?>> list = new ArrayList<>();
        for (Object person : dataList) {
            List<Object> data = new ArrayList<>();
            for (String fieldName : fileList) {
                if (person instanceof Map) {
                    data.add(((Map) person).get(fieldName));
                } else {
                    /**通过反射根据需要显示的字段，获取对应的属性值*/
                    data.add(getFieldValue(fieldName, person));
                }
            }
            list.add(data);
        }
        return list;
    }

    /**
     * 根据传入的字段获取对应的get方法，如name,对应的getName方法
     *
     * @param fieldName 字段名
     * @param person    对象
     * @return
     */
    private static Object getFieldValue(String fieldName, Object person) {
        try {
            String firstLetter = fieldName.substring(0, 1).toUpperCase();
            String getter = "get" + firstLetter + fieldName.substring(1);
            Method method = person.getClass().getMethod(getter);
            return method.invoke(person);
        } catch (Exception e) {
            return null;
        }
    }


    /**
     * 按模板写文件
     *
     * @param fileName
     * @param headClazz 表头模板
     * @param data      数据
     */
    public static <T> void write(HttpServletResponse response, String fileName, String sheetName, Class<T> headClazz, List<T> data) {
        try {
            if (StringUtils.isEmpty(sheetName)) {
                sheetName = DEFAULT_SHEET_NAME;
            }

            OutputStream outputStream = formatResponse(fileName, response);
            ExcelWriterSheetBuilder builder = EasyExcel.write(outputStream, headClazz).excelType(ExcelTypeEnum.XLSX)
                    .sheet()
                    .sheetName(sheetName);

            customCellStyle(builder, null, headClazz, data);
        } catch (Exception e) {
            log.error("export error", e);
            throw new RuntimeException(e.getMessage());
        }
    }

    /**
     * 多个sheet页的数据链式写入
     * ExcelUtil.writeWithSheets(outputStream)
     * .writeModel(ExcelModel.class, excelModelList, "sheetName1")
     * .write(headData, data,"sheetName2")
     * .finish();
     *
     * @param response
     */
    public static <T> void writeWithSheets(HttpServletResponse response, String fileName, List<MultiSheetsData<T>> multiSheetsDatas) {
        try {
            OutputStream outputStream = formatResponse(fileName, response);
            EasyExcelWriterFactory easyExcelWriterFactory = new EasyExcelWriterFactory(outputStream);

            multiSheetsDatas.forEach(multiSheetsData -> {
                easyExcelWriterFactory.writeModel(multiSheetsData.getHeadClazz(), multiSheetsData.getData(), multiSheetsData.getSheetName());
            });
            easyExcelWriterFactory.finish();
        } catch (IOException e) {
            log.error("writeWithSheets error", e);
            throw new RuntimeException(e.getMessage());
        }
    }

    /**
     * 格式化输出流，设置响应头，解决文件名乱码问题
     *
     * @param response
     * @return
     * @throws IOException
     */
    public static OutputStream formatResponse(String fileName, HttpServletResponse response) throws IOException {
        fileName = URLEncoder.encode(fileName + "_" + DateUtil.dateTimeNow() + ExcelTypeEnum.XLSX.getValue(), StandardCharsets.UTF_8.name());
        // xls
        // response.setContentType("application/vnd.ms-excel");
        // xlsx
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding(StandardCharsets.UTF_8.name());
        response.setHeader("FileName", fileName);
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName);

        return response.getOutputStream();
    }

    /**
     * 头的策略
     *
     * @return
     */
    public static WriteCellStyle getHeadWriteCellStyle(boolean isRequiredField) {
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        // 背景设置颜色
        // headWriteCellStyle.setFillForegroundColor(IndexedColors.ORANGE.index);
        headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        headWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // 边框
        headWriteCellStyle.setBorderBottom(BorderStyle.THIN);
        headWriteCellStyle.setBorderLeft(BorderStyle.THIN);
        headWriteCellStyle.setBorderRight(BorderStyle.THIN);
        headWriteCellStyle.setBorderTop(BorderStyle.THIN);
        // 自动换行
        headWriteCellStyle.setWrapped(true);
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setBold(true);
        headWriteFont.setFontName("宋体");
        headWriteFont.setFontHeightInPoints((short) 12);
        if (isRequiredField) {
            headWriteFont.setColor(Font.COLOR_RED);
        }
        headWriteCellStyle.setWriteFont(headWriteFont);
        return headWriteCellStyle;
    }

    /**
     * 内容的策略
     *
     * @return
     */
    public static WriteCellStyle getContentWriteCellStyle() {
        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.头默认了 FillPatternType所以可以不指定
        contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        // 背景绿色
        contentWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        contentWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 边框
        contentWriteCellStyle.setBorderBottom(BorderStyle.THIN);
        contentWriteCellStyle.setBorderLeft(BorderStyle.THIN);
        contentWriteCellStyle.setBorderRight(BorderStyle.THIN);
        contentWriteCellStyle.setBorderTop(BorderStyle.THIN);
        // 自动换行
        contentWriteCellStyle.setWrapped(true);
        // 文字
        WriteFont contentWriteFont = new WriteFont();
        // 字体大小
        contentWriteFont.setFontHeightInPoints((short) 12);
        contentWriteFont.setFontName("宋体");
        contentWriteCellStyle.setWriteFont(contentWriteFont);
        return contentWriteCellStyle;
    }

    /**
     * 获取样式处理器列表（指定列宽策略）
     * 如果不存在列宽策略则指定默认列宽策略
     *
     * @param writeHandlers 自定义样式处理器列表，可以为null
     * @param headClass     表头类，可以为null
     * @return 样式处理器列表
     */
    public static List<WriteHandler> getWriteHandlers(List<WriteHandler> writeHandlers, Class headClass) {

        List<WriteHandler> res = new ArrayList<>();
        // 默认表头样式
        res.addAll(getDefaultStyle(getRequiredFieldNames(headClass)));

        if (CollectionUtils.isEmpty(writeHandlers)) {
            res.add(DEFAULT_COLUMN_WIDTH_STYLE_STRATEGY);
            return res;
        }

        res.addAll(writeHandlers);

        if (writeHandlers.stream().anyMatch(t -> t instanceof AbstractColumnWidthStyleStrategy)) {
            return res;
        }

        res.add(DEFAULT_COLUMN_WIDTH_STYLE_STRATEGY);

        return res;
    }

    private static Set<String> getRequiredFieldNames(Class headClass) {
        if (headClass == null) {
            return null;
        }

        Field[] fields = ReflectUtil.getFields(headClass, t ->
                t.getAnnotation(ExcelRequired.class) != null);
        if (fields == null || fields.length == 0) {
            return null;
        }

        Set<String> result = new HashSet<>();
        for (Field field : fields) {
            WriteCellStyle headWriteCellStyle = new WriteCellStyle();

            WriteFont headWriteFont = new WriteFont();
            headWriteCellStyle.setWriteFont(headWriteFont);
            result.add(field.getName());
        }

        return result;
    }

    /**
     * 获取默认样式处理器
     * 根据必填字段名称创建默认的样式处理器
     *
     * @param requiredFiledNames 必填字段名称集合，不能为null
     * @return 样式处理器列表
     */
    private static List<WriteHandler> getDefaultStyle(Set<String> requiredFiledNames) {

        List<WriteHandler> handlerList = new ArrayList<>();

        // 内容的策略
        handlerList.add(new ExcelHorizontalCellStyleStrategy(
                getHeadWriteCellStyle(false),
                getHeadWriteCellStyle(true),
                getContentWriteCellStyle(), requiredFiledNames));

        handlerList.add(new SimpleColumnWidthStyleStrategy(DEFAULT_COLUMN_WIDTH));
        handlerList.add(new SimpleRowHeightStyleStrategy(DEFAULT_HEADER_HEIGHT, DEFAULT_CONTENT_HEIGHT));

        return handlerList;
    }

}
