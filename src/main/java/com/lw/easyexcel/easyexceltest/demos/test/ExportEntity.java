package com.lw.easyexcel.easyexceltest.demos.test;

import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * @Description:
 * @ClassName: ExportEntity
 * @Author: lw
 * @Date: 2024/7/5 16:06
 * @Version: 1.0
 */
@Data
public class ExportEntity {

    /**
     * 导出参数
     */
    private List<ExportParam> exportParams;

    /**
     * 平铺后的数据
     */
    private List<Map<String,Object>> data;

    /**
     * 每列的合并行数组
     */
    private Map<String, List<Integer[]>> mergeRowMap;


}
