package com.example.exceldemo.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.column.AbstractColumnWidthStyleStrategy;
import com.example.exceldemo.handler.ExcelMergeHandler;
import com.example.exceldemo.model.StudentExportVo;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

@Slf4j
@RestController
public class StudentController {

    @PostMapping("/exportStudent")
    public void exportStudent(@RequestBody String dynamicTitle, HttpServletResponse response) {
        List<StudentExportVo> list = this.getStudentExportVos();
        try {
            String fileName = "学生信息";
            fileName = URLEncoder.encode(fileName, "UTF-8");
            response.setContentType("application/json;charset=utf-8");
            response.setCharacterEncoding("utf-8");
            response.addHeader("Pragma", "no-cache");
            response.addHeader("Cache-Control", "no-cache");
            response.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xlsx");
            ServletOutputStream output = response.getOutputStream();
            // 需要合并的列
            int[] mergeColumnIndex = {0};
            // 从第二行后开始合并
            int mergeRowIndex = 2;
            // 设置动态标题
            List<List<String>> headers = this.getHeaders("学生信息" + dynamicTitle);
            // 头的策略
            WriteCellStyle headWriteCellStyle = new WriteCellStyle();
            // 背景设置为白色 垂直居中 水平居中 字号20
            headWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
            headWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
            WriteFont headWriteFont = new WriteFont();
            headWriteFont.setFontHeightInPoints((short) 20);
            headWriteCellStyle.setWriteFont(headWriteFont);
            // 内容的策略
            WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
            // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.头默认了 FillPatternType所以可以不指定
            // contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
            // 背景白色
            // contentWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
            // 设置 自动换行
            contentWriteCellStyle.setWrapped(true);
            // 设置 垂直居中
            contentWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            // 设置 水平居中
            contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
            // 字体大小20
            WriteFont contentWriteFont = new WriteFont();
            contentWriteFont.setFontHeightInPoints((short) 20);
            contentWriteCellStyle.setWriteFont(contentWriteFont);
            // 这个策略是 头是头的样式 内容是内容的样式 其他的策略可以自己实现
            HorizontalCellStyleStrategy horizontalCellStyleStrategy = new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);

            EasyExcel.write(output, StudentExportVo.class)
                    .sheet("学生信息")
                    .head(headers)
                    .registerWriteHandler(new ExcelMergeHandler(mergeRowIndex, mergeColumnIndex))
                    .registerWriteHandler(horizontalCellStyleStrategy)
                    // .registerWriteHandler(new SimpleColumnWidthStyleStrategy(30))
                    .registerWriteHandler(new AbstractColumnWidthStyleStrategy() {
                        @Override
                        protected void setColumnWidth(WriteSheetHolder writeSheetHolder, List<WriteCellData<?>> list, Cell cell, Head head, Integer integer, Boolean aBoolean) {
                            Sheet sheet = writeSheetHolder.getSheet();
                            int columnIndex = cell.getColumnIndex();
                            if (columnIndex == 5) {
                                // 列宽100
                                sheet.setColumnWidth(columnIndex, 10000);
                            } else {
                                // 列宽50
                                sheet.setColumnWidth(columnIndex, 5000);
                            }
                            // 行高40
                            sheet.setDefaultRowHeight((short) 4000);
                        }
                    })
                    .doWrite(list);
            output.flush();
        } catch (IOException e) {
            log.error(e.getMessage(), e);
        }
    }

    private List<List<String>> getHeaders(String dynamicTitle) {
        List<List<String>> headers = new ArrayList<>();

        List<String> schoolHead = new ArrayList<>();
        schoolHead.add(dynamicTitle);
        schoolHead.add("学校");
        headers.add(schoolHead);

        List<String> nameHead = new ArrayList<>();
        nameHead.add(dynamicTitle);
        nameHead.add("姓名");
        headers.add(nameHead);

        List<String> sexHead = new ArrayList<>();
        sexHead.add(dynamicTitle);
        sexHead.add("性别");
        headers.add(sexHead);

        List<String> ageHead = new ArrayList<>();
        ageHead.add(dynamicTitle);
        ageHead.add("年龄");
        headers.add(ageHead);

        return headers;
    }

    private List<StudentExportVo> getStudentExportVos() {
        List<StudentExportVo> list = new ArrayList<>();

        StudentExportVo v1 = new StudentExportVo();
        v1.setSchool("北京大学");
        v1.setName("张三");
        v1.setSex("男");
        v1.setAge("20");
        list.add(v1);

        StudentExportVo v2 = new StudentExportVo();
        v2.setSchool("北京大学");
        v2.setName("李四");
        v2.setSex("男");
        v2.setAge("22");
        list.add(v2);

        StudentExportVo v3 = new StudentExportVo();
        v3.setSchool("北京大学");
        v3.setName("王五");
        v3.setSex("女");
        v3.setAge("22");
        list.add(v3);

        StudentExportVo v4 = new StudentExportVo();
        v4.setSchool("清华大学");
        v4.setName("赵六");
        v4.setSex("女");
        v4.setAge("21");
        list.add(v4);

        StudentExportVo v5 = new StudentExportVo();
        v5.setSchool("武汉大学");
        v5.setName("王强");
        v5.setSex("男");
        v5.setAge("24");
        list.add(v5);

        StudentExportVo v6 = new StudentExportVo();
        v6.setSchool("武汉大学");
        v6.setName("赵燕");
        v6.setSex("女");
        v6.setAge("21");
        list.add(v6);

        StudentExportVo v7 = new StudentExportVo();
        v7.setSchool("厦门大学");
        v7.setName("陆仟");
        v7.setSex("女");
        v7.setAge("21");
        list.add(v7);

        return list;
    }

}
