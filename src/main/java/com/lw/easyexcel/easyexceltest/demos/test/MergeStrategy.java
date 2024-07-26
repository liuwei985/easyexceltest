package com.lw.easyexcel.easyexceltest.demos.test;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.merge.AbstractMergeStrategy;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.util.CollectionUtils;

import java.util.List;

public class MergeStrategy extends AbstractMergeStrategy {

    private ExportEntity exportEntity;

    public MergeStrategy(ExportEntity exportEntity) {
        this.exportEntity = exportEntity;
    }


    @Override
    protected void merge(Sheet sheet, Cell cell, Head head, Integer relativeRowIndex) {
        //mergeMap需要加上head的行数
        String fieldName = head.getHeadNameList().get(0);
        List<Integer[]> mergeRow = exportEntity.getMergeRowMap().get(fieldName);
        if (CollectionUtils.isEmpty(mergeRow)) return;
        int currentColumnIndex = cell.getColumnIndex();
        int currentRowIndex = cell.getRowIndex();
        //表头所占行数
        int headHeight = head.getHeadNameList().size();
        for (int i = 0; i < mergeRow.size(); i++) {
            if (currentRowIndex == mergeRow.get(i)[1] + headHeight - 1) {
                if (mergeRow.get(i)[1] + headHeight - 1 == mergeRow.get(i)[0] + headHeight) {
                    return;
                }
                sheet.addMergedRegion(new CellRangeAddress(mergeRow.get(i)[0] + headHeight, mergeRow.get(i)[1] + headHeight - 1,
                        currentColumnIndex, currentColumnIndex));
            }
        }

    }

}
