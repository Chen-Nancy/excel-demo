package com.example.exceldemo.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentRowHeight;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

@NoArgsConstructor
@AllArgsConstructor
@Data
@ContentRowHeight(30)
@HeadRowHeight(40)
@ColumnWidth(25)
public class StudentExportVo implements Serializable {
    private static final long serialVersionUID = -5809782578272943999L;

    // @ContentLoopMerge(eachRow = 3)
    @ExcelProperty(value = {"学生信息", "学校"}, order = 1)
    // @ExcelProperty(value = "学校", order = 1)
    private String school;

    @ExcelProperty(value = {"学生信息", "姓名"}, order = 2)
    // @ExcelProperty(value = "姓名", order = 2)
    private String name;

    @ExcelProperty(value = {"学生信息", "性别"}, order = 3)
    // @ExcelProperty(value = "性别", order = 3)
    private String sex;

    @ExcelProperty(value = {"学生信息", "年龄"}, order = 4)
    // @ExcelProperty(value = "年龄", order = 4)
    private String age;
}
