package com.lw.easyexcel.easyexceltest.demos.test;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class HeadHandler implements CellWriteHandler {

    private ExportEntity exportEntity;


    public HeadHandler(ExportEntity exportEntity) {
        this.exportEntity = exportEntity;
    }

    @Override
    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer integer, Integer integer1, Boolean aBoolean) {

    }

    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell, Head head, Integer integer, Boolean aBoolean) {

    }

    @Override
    public void afterCellDataConverted(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, WriteCellData cellData, Cell cell, Head head, Integer integer, Boolean aBoolean) {

    }

    //    @Override
    //    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> list, Cell cell, Head head, Integer integer, Boolean aBoolean) {
    //        if (!aBoolean) return;
    //        List<String> headNameList = head.getHeadNameList();
    //        Map<String, String> collect = exportEntity.getExportParams().stream().collect(Collectors.toMap(ExportParam::getFieldName, ExportParam::getColumnName));
    //        cell.setCellValue(collect.get(headNameList.get(0)));
    //    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<WriteCellData<?>> list, Cell cell, Head head, Integer integer, Boolean aBoolean) {
        if (!aBoolean) return;
        List<String> headNameList = head.getHeadNameList();
        Map<String, String> collect = exportEntity.getExportParams().stream().collect(Collectors.toMap(ExportParam::getFieldName, ExportParam::getColumnName));
        cell.setCellValue(collect.get(headNameList.get(0)));
    }
}

