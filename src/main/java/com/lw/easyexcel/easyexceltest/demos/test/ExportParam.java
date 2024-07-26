package com.lw.easyexcel.easyexceltest.demos.test;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * @Description:
 * @ClassName: ExportParam
 * @Author: lw
 * @Date: 2024/7/5 16:09
 * @Version: 1.0
 */
@Data
@Accessors(chain = true)
public class ExportParam {
    private String fieldName;
    private String columnName;
    private Integer order;
}
