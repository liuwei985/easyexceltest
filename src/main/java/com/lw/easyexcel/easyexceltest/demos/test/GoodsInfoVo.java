package com.lw.easyexcel.easyexceltest.demos.test;

import com.lw.easyexcel.easyexceltest.annotation.MyExcelProperty;
import lombok.Data;

import java.math.BigDecimal;

/**
 * @Description:
 * @ClassName: GoodsInfoVo
 * @Author: lw
 * @Date: 2024/7/5 17:25
 * @Version: 1.0
 */
@Data
public class GoodsInfoVo {

    @MyExcelProperty(name = "商品ID",index = 1)
    private Long goodsId;

    @MyExcelProperty(name = "商品名称",index = 2)
    private String goodsName;

    @MyExcelProperty(name = "商品单价",index = 3)
    private BigDecimal price;

}
