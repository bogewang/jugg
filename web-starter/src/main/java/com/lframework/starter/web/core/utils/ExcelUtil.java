package com.lframework.starter.web.core.utils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.handler.WriteHandler;
import com.google.common.collect.Lists;
import com.lframework.starter.common.exceptions.impl.DefaultSysException;
import com.lframework.starter.common.utils.CollectionUtil;
import com.lframework.starter.common.utils.FileUtil;
import com.lframework.starter.web.core.components.excel.ExcelHorizontalCellStyleStrategy;
import com.lframework.starter.web.core.components.excel.ExcelModel;
import com.lframework.starter.web.core.components.excel.ExcelMultipartWriterBuilder;
import com.lframework.starter.web.core.components.excel.ExcelMultipartWriterSheetBuilder;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

import static com.lframework.starter.web.core.utils.EasyExcelUtils.*;

/**
 * Excel工具类
 * 提供Excel文件导入导出功能，基于EasyExcel实现
 * 包括Excel读取、写入、样式设置、数据验证等功能
 *
 * @author lframework@163.com
 */
@Slf4j
public class ExcelUtil {

    /**
     * 读取Excel文件
     * 使用EasyExcel读取Excel文件并返回读取构建器
     *
     * @param file     Excel文件，不能为null
     * @param clazz    目标类型，不能为null
     * @param listener 读取监听器，不能为null
     * @param <T>      目标类型泛型
     * @return Excel读取构建器
     * @throws DefaultSysException 当文件读取失败时抛出
     */
    public static <T> ExcelReaderBuilder read(MultipartFile file, Class<T> clazz, ReadListener<T> listener) {
        return EasyExcelUtils.read(file, clazz, listener);
    }

    /**
     * 导出Xlsx至Response
     *
     * @param sheetName
     * @param clazz
     * @param <T>
     */
    public static <T extends ExcelModel> void export(String sheetName, Class<T> clazz) {
        HttpServletResponse response = ResponseUtil.getResponse();
        EasyExcelUtils.write(response, sheetName, clazz, Lists.newArrayList());
    }

    /**
     * 导出Xlsx至Response
     *
     * @param sheetName
     * @param clazz
     * @param datas
     * @param <T>
     */
    public static <T extends ExcelModel> void export(String sheetName, Class<T> clazz, List<T> datas) {
        HttpServletResponse response = ResponseUtil.getResponse();
        EasyExcelUtils.write(response, sheetName, clazz, datas);
    }

    /**
     * 导出Xlsx至Response，动态表头
     *
     * @param sheetName
     * @param headerMap
     * @param datas
     * @param <T>
     */
    public static <T extends ExcelModel> void export(String sheetName, Map<String, String> headerMap, List<T> datas) {
        HttpServletResponse response = ResponseUtil.getResponse();
        EasyExcelUtils.write(response, sheetName, headerMap, datas);
    }

    /**
     * 导出Xlsx至Response，动态表头
     *
     * @param sheetName
     * @param headerMap
     * @param datas
     * @param <T>
     */
    public static <T extends ExcelModel> void exportNoModel(String sheetName, Map<String, String> headerMap, List<Map<String, String>> datas) {
        HttpServletResponse response = ResponseUtil.getResponse();
        EasyExcelUtils.writeNoModel(response, sheetName, headerMap, datas);
    }

    /**
     * 导出XLSX文件到响应（使用空数据）
     * 使用指定文件名，导出空的Excel文件
     *
     * @param fileName  文件名，不能为null
     * @param sheetName 工作表名称，不能为null
     * @param clazz     Excel模型类，不能为null
     * @param <T>       Excel模型类型
     */
    public static <T extends ExcelModel> void export(String fileName, String sheetName, Class<T> clazz) {
        // export(fileName, sheetName, clazz, CollectionUtil.emptyList(), null);
        HttpServletResponse response = ResponseUtil.getResponse();
        EasyExcelUtils.write(response, fileName, sheetName, clazz, Lists.newArrayList());
    }

    /**
     * 导出XLSX文件到响应（使用默认样式处理器）
     * 使用指定文件名和数据，导出Excel文件
     *
     * @param fileName  文件名，不能为null
     * @param sheetName 工作表名称，不能为null
     * @param clazz     Excel模型类，不能为null
     * @param datas     数据列表，不能为null
     * @param <T>       Excel模型类型
     */
    public static <T extends ExcelModel> void export(String fileName, String sheetName, Class<T> clazz, List<T> datas) {
        // export(fileName, sheetName, clazz, datas, null);
        HttpServletResponse response = ResponseUtil.getResponse();
        EasyExcelUtils.write(response, fileName, sheetName, clazz, datas);
    }

    /**
     * 导出XLSX文件到响应（完整参数）
     * 使用指定文件名、数据和样式处理器，导出Excel文件
     *
     * @param fileName      文件名，不能为null
     * @param sheetName     工作表名称，不能为null
     * @param clazz         Excel模型类，不能为null
     * @param datas         数据列表，不能为null
     * @param writeHandlers 样式处理器列表，可以为null
     * @param <T>           Excel模型类型
     * @throws DefaultSysException 当文件写入失败时抛出
     */
    public static <T extends ExcelModel> void export(String fileName, String sheetName,
                                                     Class<T> clazz, List<T> datas, List<WriteHandler> writeHandlers) {
        try {
            HttpServletResponse response = ResponseUtil.getResponse();
            OutputStream os = EasyExcelUtils.formatResponse(fileName, response);

            exportExcel(os, sheetName, ExcelTypeEnum.XLSX, clazz, datas, writeHandlers);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            throw new DefaultSysException("导出异常");
        }
    }

    /**
     * 导出Xlsx至文件
     *
     * @param sheetName
     * @param clazz
     * @param <T>
     */
    public static <T extends ExcelModel> void export(File file, String sheetName, Class<T> clazz) {
        export(file, sheetName, clazz, CollectionUtil.emptyList(), null);
    }

    /**
     * 导出Xlsx至文件
     *
     * @param sheetName
     * @param clazz
     * @param datas
     * @param <T>
     */
    public static <T extends ExcelModel> void export(File file, String sheetName, Class<T> clazz, List<T> datas) {
        export(file, sheetName, clazz, datas, null);
    }

    /**
     * 导出Xlsx至文件
     *
     * @param sheetName
     * @param clazz
     * @param datas
     * @param writeHandlers
     * @param <T>
     */
    public static <T extends ExcelModel> void export(File file, String sheetName, Class<T> clazz,
                                                     List<T> datas, List<WriteHandler> writeHandlers) {
        exportExcel(FileUtil.getOutputStream(file), sheetName, ExcelTypeEnum.XLSX, clazz, datas, writeHandlers);
    }

    /**
     * 分段导出Excel
     *
     * @param os
     * @param sheetName
     * @param excelType
     * @param head
     * @param <T>
     */
    private static <T extends ExcelModel> ExcelMultipartWriterSheetBuilder multipartExportExcel(
            OutputStream os,
            String sheetName, ExcelTypeEnum excelType, List<String> head) {

        List<List<String>> headWrapper = new ArrayList<>();
        if (!CollectionUtil.isEmpty(head)) {
            for (String s : head) {
                headWrapper.add(Collections.singletonList(s));
            }
        }

        ExcelMultipartWriterSheetBuilder builder = new ExcelMultipartWriterBuilder().file(os)
                .excelType(excelType)
                .useDefaultStyle(false).head(headWrapper).sheet(sheetName);
        List<WriteHandler> writeHandlers = getWriteHandlers();
        writeHandlers.forEach(builder::registerWriteHandler);

        return builder;
    }

    /**
     * 分段导出XLSX文件到响应（使用默认文件名）
     * 使用工作表名作为文件名，创建分段导出构建器
     *
     * @param sheetName 工作表名称，不能为null
     * @param clazz     Excel模型类，不能为null
     * @param <T>       Excel模型类型
     * @return 分段导出构建器
     */
    public static <T extends ExcelModel> ExcelMultipartWriterSheetBuilder multipartExportXlsx(
            String sheetName,
            Class<T> clazz) {

        return multipartExportXlsx(sheetName, sheetName, clazz);
    }

    /**
     * 分段导出XLSX文件到响应（指定文件名）
     * 使用指定文件名，创建分段导出构建器
     *
     * @param fileName  文件名，不能为null
     * @param sheetName 工作表名称，不能为null
     * @param clazz     Excel模型类，不能为null
     * @param <T>       Excel模型类型
     * @return 分段导出构建器
     * @throws DefaultSysException 当文件写入失败时抛出
     */
    public static <T extends ExcelModel> ExcelMultipartWriterSheetBuilder multipartExportXlsx(String fileName,
                                                                                              String sheetName, Class<T> clazz) {
        try {
            HttpServletResponse response = ResponseUtil.getResponse();
            OutputStream os = EasyExcelUtils.formatResponse(fileName, response);

            return multipartExportExcel(os, sheetName, ExcelTypeEnum.XLSX, clazz);
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            throw new DefaultSysException("导出异常");
        }
    }

    /**
     * 分段导出Xlsx至文件
     *
     * @param sheetName
     * @param clazz
     * @param <T>
     */
    public static <T extends ExcelModel> ExcelMultipartWriterSheetBuilder multipartExportXlsx(File file,
                                                                                              String sheetName, Class<T> clazz) {

        return multipartExportExcel(FileUtil.getOutputStream(file), sheetName, ExcelTypeEnum.XLSX, clazz);
    }

    /**
     * 导出Excel
     *
     * @param os
     * @param sheetName
     * @param excelType
     * @param clazz
     * @param datas
     * @param writeHandlers
     * @param <T>
     */
    private static <T extends ExcelModel> void exportExcel(OutputStream os, String sheetName,
                                                           ExcelTypeEnum excelType, Class<T> clazz, List<T> datas, List<WriteHandler> writeHandlers) {

        ExcelMultipartWriterSheetBuilder builder = new ExcelMultipartWriterBuilder().file(os)
                .excelType(excelType)
                .useDefaultStyle(false).head(clazz).sheet(sheetName);
        writeHandlers = EasyExcelUtils.getWriteHandlers(writeHandlers, clazz);

        writeHandlers.forEach(builder::registerWriteHandler);

        builder.doWrite(datas);
        builder.finish();
    }

    /**
     * 分段导出Excel
     *
     * @param os
     * @param sheetName
     * @param excelType
     * @param clazz
     * @param <T>
     */
    private static <T extends ExcelModel> ExcelMultipartWriterSheetBuilder multipartExportExcel(
            OutputStream os, String sheetName, ExcelTypeEnum excelType, Class<T> clazz) {

        ExcelMultipartWriterSheetBuilder builder = new ExcelMultipartWriterBuilder().file(os)
                .excelType(excelType)
                .useDefaultStyle(false).head(clazz).sheet(sheetName);

        List<WriteHandler> writeHandlers = EasyExcelUtils.getWriteHandlers(null, clazz);
        writeHandlers.forEach(builder::registerWriteHandler);

        return builder;
    }

    /**
     * 获取默认样式处理器列表
     * 获取包含默认样式的WriteHandler列表
     *
     * @return 样式处理器列表
     */
    public static List<WriteHandler> getWriteHandlers() {

        return EasyExcelUtils.getWriteHandlers(null, null);
    }

    /**
     * 多sheet导出
     * @param fileName
     * @param multiSheetsDatas
     */
    public static <T extends ExcelModel> void writeWithSheets(String fileName,  List<MultiSheetsData<T>> multiSheetsDatas) {
        HttpServletResponse response = ResponseUtil.getResponse();
        EasyExcelUtils.writeWithSheets(response, fileName, multiSheetsDatas);
    }

}
