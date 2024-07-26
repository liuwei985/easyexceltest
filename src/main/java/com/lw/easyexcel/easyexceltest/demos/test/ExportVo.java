package com.lw.easyexcel.easyexceltest.demos.test;

import com.lw.easyexcel.easyexceltest.annotation.MyExcelCollection;
import com.lw.easyexcel.easyexceltest.annotation.MyExcelEntity;
import com.lw.easyexcel.easyexceltest.annotation.MyExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import java.util.List;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExportVo {
    @MyExcelProperty(name = "小组名称", index = 1)
    private String groupName;
    @MyExcelProperty(name = "小组口号", index = 2)
    private String groupSlogan;
    private String groupType;
    @MyExcelEntity(name = "房子")
    private House house;
    @MyExcelCollection(name = "组员信息")
    private List<ProjectGroupExcelVO> groupUsers;


    @Builder
    @Data
    public static class ProjectGroupExcelVO {
        @MyExcelProperty(name = "组员姓名")
        private String name;
        @MyExcelProperty(name = "组员电话")
        private String phone;
        @MyExcelProperty(name = "年龄")
        private Integer age;
        @MyExcelCollection(name = "车辆信息")
        private List<Vehicle> vehicle;

    }

    @Builder
    @Data
    public static class Company {
        @MyExcelProperty(name = "公司名称")
        private String companyName;
        @MyExcelProperty(name = "公司地址")
        private String address;

    }


    @Builder
    @Data
    public static class Vehicle {
        @MyExcelProperty(name = "车牌")
        private String plateNo;
        @MyExcelProperty(name = "品牌")
        private String brand;

    }

    @Data
    @Accessors(chain = true)
    public static class House {
        @MyExcelProperty(name = "住址名称")
        private String addressName;
        @MyExcelCollection(name = "房间")
        private List<Room> roomList;
    }

    @Data
    @Accessors(chain = true)
    public static class Room {
        @MyExcelProperty(name = "房间名")
        private String name;
        @MyExcelProperty(name = "大小")
        private Integer square;
    }
}

