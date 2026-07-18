package com.lframework.starter.web.core.utils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.lframework.starter.web.core.components.excel.ExcelModel;

import java.io.File;
import java.io.OutputStream;
import java.util.List;

import static com.lframework.starter.web.core.utils.EasyExcelUtils.customCellStyle;
import static com.lframework.starter.web.core.utils.EasyExcelUtils.customCellStyleNoWrite;

public class EasyExcelWriterFactory {

    private int sheetNo = 0;
    private ExcelWriter excelWriter = null;

    public EasyExcelWriterFactory(OutputStream outputStream) {
        excelWriter = EasyExcel.write(outputStream).build();
    }

    public EasyExcelWriterFactory(File file) {
        excelWriter = EasyExcel.write(file).build();
    }

    public EasyExcelWriterFactory(String filePath) {
        excelWriter = EasyExcel.write(filePath).build();
    }

    /**
     * 链式模板表头写入
     *
     * @param headClazz 表头格式
     * @param data      数据 List<ExcelModel> 或者List<List<Object>>
     * @return
     */
    public EasyExcelWriterFactory writeModel(Class headClazz, List data) {
        excelWriter.write(data, EasyExcel.writerSheet(this.sheetNo++).head(headClazz).build());
        return this;
    }


    /**
     * 链式模板表头写入
     *
     * @param headClazz 表头格式
     * @param data      数据 List<ExcelModel> 或者List<List<Object>>
     * @return
     */
    public <T> EasyExcelWriterFactory writeModel(Class<T> headClazz, List<T> data, String sheetName) {
        ExcelWriterSheetBuilder sheetBuilder = EasyExcel.writerSheet(this.sheetNo++, sheetName).head(headClazz);
        customCellStyleNoWrite(sheetBuilder, null, headClazz, data);

        WriteSheet writeSheet = sheetBuilder.build();
        excelWriter.write(data, writeSheet);
        return this;
    }

    /**
     * 链式自定义表头写入
     *
     * @param head
     * @param data      数据 List<ExcelModel> 或者List<List<Object>>
     * @param sheetName
     * @return
     */
    public EasyExcelWriterFactory write(List<List<String>> head, List data, String sheetName) {
        excelWriter.write(data, EasyExcel.writerSheet(this.sheetNo++, sheetName).head(head).build());
        return this;
    }

    /**
     * 使用此类结束后，一定要关闭流
     */
    public void finish() {
        excelWriter.finish();
    }
}

