package com.lw.easyexcel.easyexceltest.demos.test;

import com.alibaba.excel.event.Order;
import com.google.common.collect.Lists;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

@Controller
public class ExcelController {


    //构造一对多嵌套数据
    public static List<Object> build() {
        List<Object> exportVos = new ArrayList<>();
        //第一条数据
        List<ExportVo.ProjectGroupExcelVO> projectGroupList = new ArrayList<>();
        ExportVo.ProjectGroupExcelVO projectGroupExcelVO1 = ExportVo.ProjectGroupExcelVO.builder()
                .age(17)
                .name("张三")
                .phone("111111111111")
                .vehicle(Lists.newArrayList(ExportVo.Vehicle.builder().brand("BYD").plateNo("京AM8888").build()
                        , ExportVo.Vehicle.builder().brand("BMW").plateNo("京AM9999").build()))
                .build();

        ExportVo.ProjectGroupExcelVO projectGroupExcelVO2 = ExportVo.ProjectGroupExcelVO.builder()
                .age(18)
                .name("李四")
                .phone("22222222222")
                .vehicle(Lists.newArrayList(ExportVo.Vehicle.builder().brand("NIO").plateNo("沪A99999").build(),
                        ExportVo.Vehicle.builder().brand("xiaopeng").plateNo("沪A88888").build()
                        , ExportVo.Vehicle.builder().brand("Benz").plateNo("沪B7777").build()
                        )
                )
                .build();


        projectGroupList.add(projectGroupExcelVO1);
        projectGroupList.add(projectGroupExcelVO2);
        ExportVo exportVo = ExportVo.builder().groupName("groupOne")
                .groupSlogan("we are family")
                .groupType("aaa")
                .house(new ExportVo.House()
                        .setAddressName("北京四合院")
                        .setRoomList(Lists.newArrayList(new ExportVo.Room().setName("主卧").setSquare(130),
                                new ExportVo.Room().setName("次卧").setSquare(100))))
                .build();
        exportVo.setGroupUsers(projectGroupList);


        //第二条数据
        List<ExportVo.ProjectGroupExcelVO> projectGroupList2 = new ArrayList<>();
        ExportVo.ProjectGroupExcelVO projectGroupExcelVO22 = ExportVo.ProjectGroupExcelVO.builder()
                .age(27)
                .name("王五")
                .phone("444444444444")
                .build();

        ExportVo.ProjectGroupExcelVO projectGroupExcelVO23 = ExportVo.ProjectGroupExcelVO.builder()
                .age(28)
                .name("老六")
                .phone("5555555555555")
                .build();

        projectGroupList2.add(projectGroupExcelVO22);
        projectGroupList2.add(projectGroupExcelVO23);
        ExportVo exportVo2 = ExportVo.builder().groupName("groupTwo")
                .groupSlogan("we are friends")
                .groupType("bbb")
                .house(new ExportVo.House().setAddressName("上海")
                        .setRoomList(Lists.newArrayList(new ExportVo.Room().setName("401"),
                                new ExportVo.Room().setName("402"))))
                .build();
        exportVo2.setGroupUsers(projectGroupList2);
        exportVos.add(exportVo);
        exportVos.add(exportVo2);
        return exportVos;
    }


    @RequestMapping("/orderList")
    public void export(HttpServletResponse response) {
        List<Object> list=new ArrayList<>();
        OrderInfoVo vo=new OrderInfoVo();
        vo.setOrderId(73878718781313L);
        vo.setCreateTime("2024-06-23");
        vo.setMoney(new BigDecimal("23.30"));

        List<GoodsInfoVo> goodsInfoVos=new ArrayList<>();
        GoodsInfoVo goodsInfoVo=new GoodsInfoVo();
        goodsInfoVo.setGoodsId(1L);
        goodsInfoVo.setGoodsName("测试商品1");
        goodsInfoVo.setPrice(new BigDecimal("10.32"));
        goodsInfoVos.add(goodsInfoVo);

        GoodsInfoVo goodsInfoVo2=new GoodsInfoVo();
        goodsInfoVo2.setGoodsId(2L);
        goodsInfoVo2.setGoodsName("测试商品2");
        goodsInfoVo2.setPrice(new BigDecimal("9.99"));
        goodsInfoVos.add(goodsInfoVo2);
        vo.setGoodsInfoVoList(goodsInfoVos);

        list.add(vo);

        OrderInfoVo vo2=new OrderInfoVo();
        vo2.setOrderId(738787187817878L);
        vo2.setCreateTime("2024-09-23");
        vo2.setMoney(new BigDecimal("199.30"));

//        List<GoodsInfoVo> goodsInfoVos2=new ArrayList<>();
//        GoodsInfoVo goodsInfoVo3=new GoodsInfoVo();
//        goodsInfoVo3.setGoodsId(12L);
//        goodsInfoVo3.setGoodsName("测试商品12");
//        goodsInfoVo3.setPrice(new BigDecimal("199.32"));
//        goodsInfoVos2.add(goodsInfoVo3);

//        GoodsInfoVo goodsInfoVo4=new GoodsInfoVo();
//        goodsInfoVo4.setGoodsId(12L);
//        goodsInfoVo4.setGoodsName("测试商品12");
//        goodsInfoVo4.setPrice(new BigDecimal("199.32"));
//        goodsInfoVos2.add(goodsInfoVo4);

//        vo2.setGoodsInfoVoList(goodsInfoVos2);

        list.add(vo2);


        //调用上述封装好的导出方法
        new MyExcelUtil().exportData(list, "data2", response);
    }


}

