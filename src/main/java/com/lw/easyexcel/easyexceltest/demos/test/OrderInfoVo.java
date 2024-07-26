package com.lw.easyexcel.easyexceltest.demos.test;

import com.lw.easyexcel.easyexceltest.annotation.MyExcelCollection;
import com.lw.easyexcel.easyexceltest.annotation.MyExcelProperty;
import lombok.Data;

import java.math.BigDecimal;
import java.util.List;

/**
 * @Description:
 * @ClassName: OrderInfoVo
 * @Author: lw
 * @Date: 2024/7/5 17:23
 * @Version: 1.0
 */
@Data
public class OrderInfoVo {

    @MyExcelProperty(name = "订单号", index = 1)
    private Long orderId;

    @MyExcelProperty(name = "金额", index = 2)
    private BigDecimal money;

    @MyExcelProperty(name = "时间", index = 3)
    private String createTime;

    @MyExcelCollection(name = "商品")
    private List<GoodsInfoVo> goodsInfoVoList;
}
