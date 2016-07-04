package com.neunn.excelutils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;

import com.neunn.excelutils.excelentity.ExcelExportSheetEntity;

/**
 * @author hjl
 * @E-mail:huangjl@neunn.com
 * @version 创建时间：2015年10月19日 下午5:52:54
 */
public class ExcelParser {

    public static final String VALUED_DELIM = "${";

    public static final String VALUED_DELIM2 = "}";

    public static final String KEY_TAG = "#";

    public static final String KEY_TAG_LOOP = "*{";

    public static final String VALUED_COLUMN = "&{";

    public static final String NUMBER = "number";

    /**
     * parse the Excel template
     * 
     */
    public static void parse(ExcelExportSheetEntity entity, Sheet sheet) {
        
        if(null == entity) {
            return  ;
        }
        
        List<List<Map<String, String>>> loopData                = entity.getManyRepeated() ;
        Map<String, List<Map<String, String>>> mergedLoopData   = entity.getCellMergedRepeated() ;
        Map<String, Map<String, String>> cellMergedOnce         = entity.getCellMergedOnce() ;
        List<Map<String, String>> manyColumnData                = entity.getManyColumn() ;
        
        
        int firstRow = sheet.getFirstRowNum() ;
        int loopRowNumber = 0 ;
        
        // 拷贝需要循环的列
        if(manyColumnData != null) {
            int lastRow = sheet.getLastRowNum() ;
            Integer[] columnIndex = findLoopColumnNumber(sheet, firstRow, lastRow) ;
            int rowNum              = columnIndex[0] ;
            int columnNum           = columnIndex[1] ;
            int copySize            = columnIndex[2] ;
            if(rowNum > 0 && columnNum > 0 && copySize > 0) {
                copyColumn(sheet, rowNum, lastRow, columnNum, copySize, manyColumnData) ;
                
                for(int i = firstRow ; i <= lastRow; i ++) {
                    Row row = sheet.getRow(i) ;
                    if(null == row) {
                        continue ;
                    }
                    for(int j = 0; j < manyColumnData.size(); j++) {
                        Map<String, String> oneData = manyColumnData.get(j) ;
                        setColumnLoopData(row, columnNum + copySize * j, copySize, oneData) ;
                    }
                }
            }
        }
        
        // 设置一次性数据
        setOnceDataValue(sheet, entity.getOnce()) ;
        
        // 设置循环值
        if(loopData != null && !loopData.isEmpty()) {
            
            for(List<Map<String, String>> loopSize : loopData) {
                int lastRow = sheet.getLastRowNum() ;
                loopRowNumber = findLoopRowNumber(sheet, firstRow, lastRow) ;
                int size = loopSize.size() ;
                if(size > 1) {
                    WorkbookUtils.copyRow(sheet, loopRowNumber, loopRowNumber + 1, size - 1);
                }
                setLoopValue(sheet, loopRowNumber, loopRowNumber + 1, loopSize) ;
                
                if(size > 1) {
                    firstRow = loopRowNumber + size - 1 ;
                }else{
                    firstRow = loopRowNumber + 1 ;
                }
                
            }
        }if (mergedLoopData != null && !mergedLoopData.isEmpty()) {
            
            List<Map<String, String>> newLoopData = new ArrayList<Map<String, String>>() ;
            Set<String> keys = mergedLoopData.keySet() ;
            if(keys != null && !keys.isEmpty()) {
                int twoDataSize = 0 ; 
                for(String key : keys) {
                    List<Map<String, String>> twoData = mergedLoopData.get(key) ;
                    twoDataSize = twoData.size() == 0 ? 1 : twoData.size() ; // 二次内容默认一条记录
                    break ;
                }
                
                int lastRow = sheet.getLastRowNum() ;
                loopRowNumber = findLoopRowNumber(sheet, firstRow, lastRow) ;
                copyMutilRow(sheet, loopRowNumber, keys.size(), twoDataSize) ;
                int MutilSize = lastRow - loopRowNumber + twoDataSize;
                
                // 对内部的循环行赋值
                int i = 0 ;
                for(String key : keys) {
                    List<Map<String, String>> twoData = mergedLoopData.get(key) ;
                    int num = findLoopRowNumberNoMerged(sheet, loopRowNumber + (loopRowNumber + twoDataSize) * i, loopRowNumber + MutilSize * (i + 1)) ;
                    setLoopValue(sheet, num, num + 1, twoData) ;
                    i++;
                }
                
                if(cellMergedOnce != null) {
                   Set<String> keysets = cellMergedOnce.keySet() ;
                   for(String key : keysets) {
                       Map<String, String> once = cellMergedOnce.get(key) ;
                       once.put("key", key) ;
                       newLoopData.add(once) ;
                   }
                   
                   setLoopValue(sheet, loopRowNumber, lastRow + twoDataSize,newLoopData) ;
                }
                
            }else{
                setLoopNullValue(sheet) ;
            }
            
        }else{
            setLoopNullValue(sheet) ;
        }
    }

    /**
     * 设置一次数据内容
     */
    public static void setOnceDataValue(Sheet sheet, Map<String, String> data) {

        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();

        Set<String> keys = data.keySet();

        for (int i = firstRow; i <= lastRow; i++) {
            Row row = sheet.getRow(i);
            short firstCellNum = row.getFirstCellNum();
            short lastCellNum = row.getLastCellNum();

            for (short j = firstCellNum; j <= lastCellNum; j++) {
                Cell cell = row.getCell(j);

                if (null == cell
                        || cell.getCellType() != XSSFCell.CELL_TYPE_STRING) {
                    continue;
                }

                String cellstr = cell.getStringCellValue();
                if (null == cellstr || "".equals(cellstr)) {
                    continue;
                }

                if (cellstr.startsWith(KEY_TAG)) {

                    if (keys.isEmpty()) {
                        cell.setCellValue("");
                        continue;
                    }

                    cellstr = cellstr.replace(KEY_TAG, "");
                    if (keys.contains(cellstr)) {
                        cell.setCellValue(data.get(cellstr));
                    }
                }
            }
        }

    }

    public static void setLoopNullValue(Sheet sheet) {

        int startRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();

        for (int i = startRow; i <= lastRow; i++) {
            Row row = sheet.getRow(i);
            short firstNum = row.getFirstCellNum();
            short lastNum = row.getLastCellNum();

            for (int j = firstNum; j <= lastNum; j++) {
                Cell cell = row.getCell(j);

                if (null == cell
                        || cell.getCellType() != XSSFCell.CELL_TYPE_STRING) {
                    continue;
                }

                String cellstr = cell.getStringCellValue();
                if (null == cellstr || "".equals(cellstr)) {
                    continue;
                }

                if (cellstr.startsWith(VALUED_DELIM)
                        && cellstr.endsWith(VALUED_DELIM2)) {
                    cell.setCellValue("");
                }
            }
        }

    }

    /**
     * 设置循环值的内容
     */
    public static void setLoopValue(Sheet sheet, int startRow, int endRow,
            List<Map<String, String>> datas) {

        int num = 1;

        int size = endRow - startRow;

        try {
            
            if(datas.isEmpty()) {
                setValue(sheet, startRow, num, new HashMap<String, String>());
            }
            
            for (Map<String, String> data : datas) {

                if (size == 1) {
                    setValue(sheet, startRow, num, data);
                } else if (size > 1) {
                    for (int i = 1; i <= size; i++) {
                        setValue(sheet, startRow + (num - 1) * size, i, data);
                    }
                }

                num++;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public static void setValue(Sheet sheet, int startRow, int num,
            Map<String, String> data) {
        Row row = sheet.getRow(startRow + num - 1);
        Set<String> keys = data.keySet();
        if (null == row) {
            return;
        }
        short firstNum = row.getFirstCellNum();
        short lastNum = row.getLastCellNum();

        for (int i = firstNum; i <= lastNum; i++) {
            Cell cell = row.getCell(i);

            if (null == cell || cell.getCellType() != XSSFCell.CELL_TYPE_STRING) {
                continue;
            }

            String cellstr = cell.getStringCellValue();
            if (null == cellstr || "".equals(cellstr)) {
                continue;
            }

            if ((cellstr.startsWith(VALUED_DELIM) && cellstr
                    .endsWith(VALUED_DELIM2))
                    || (cellstr.startsWith(KEY_TAG_LOOP) && cellstr
                            .endsWith(VALUED_DELIM2))) {

                if (keys.isEmpty()) {
                    cell.setCellValue("");
                    continue;
                }

                cellstr = cellstr.replace(VALUED_DELIM, "")
                        .replace(VALUED_DELIM2, "").replace(KEY_TAG_LOOP, "");

                if (keys.contains(cellstr)) {
                    cell.setCellValue(data.get(cellstr));
                }

                if (NUMBER.equals(cellstr)) {
                    cell.setCellValue(num);
                }
                
                if (keys.isEmpty()) {
                    cell.setCellValue("");
                }
            }

            if (cellstr.contains(VALUED_DELIM)
                    && cellstr.contains(VALUED_DELIM2)) {
                int startIndex = cellstr.indexOf(VALUED_DELIM)
                        + VALUED_DELIM.length();
                int endIndex = cellstr.indexOf(VALUED_DELIM2)
                        + VALUED_DELIM2.length();
                String start = cellstr.substring(0,
                        startIndex - VALUED_DELIM.length());
                String value = cellstr.substring(startIndex, endIndex
                        - VALUED_DELIM2.length());
                String end = cellstr.substring(endIndex);

                if (keys.contains(value)) {
                    cell.setCellValue(new StringBuilder(start)
                            .append(data.get(value)).append(end).toString());
                }
                
                if (keys.isEmpty()) {
                    cell.setCellValue("");
                }
            }
        }

    }

    /**
     * 查询循环开始行数
     */
    public static Integer findLoopRowNumber(Sheet sheet, int firstRow,
            int lastRow) {

        for (int i = firstRow; i <= lastRow; i++) {
            Row row = sheet.getRow(i);
            short firstCellNum = row.getFirstCellNum();
            short lastCellNum = row.getLastCellNum();

            for (short j = firstCellNum; j <= lastCellNum; j++) {
                Cell cell = row.getCell(j);

                if (null == cell
                        || cell.getCellType() != XSSFCell.CELL_TYPE_STRING) {
                    continue;
                }

                String cellstr = cell.getStringCellValue();
                if (null == cellstr || "".equals(cellstr)) {
                    continue;
                }

                if (cellstr.startsWith(VALUED_DELIM)
                        && cellstr.endsWith(VALUED_DELIM2)) {
                    return i;
                }
            }
        }

        return 0;
    }

    /**
     * 返回行与列的index
     */
    public static Integer[] findLoopColumnNumber(Sheet sheet, int firstRow,
            int lastRow) {

        Integer[] rowAndColumnAndSize = { 0, 0, 0 };

        boolean isFirst = true;
        int size = 0;

        for (int i = firstRow; i <= lastRow; i++) {
            Row row = sheet.getRow(i);
            short firstCellNum = row.getFirstCellNum();
            short lastCellNum = row.getLastCellNum();

            if (!isFirst) {
                rowAndColumnAndSize[2] = size;
                return rowAndColumnAndSize;
            }

            for (int j = firstCellNum; j <= lastCellNum; j++) {
                Cell cell = row.getCell(j);

                if (null == cell
                        || cell.getCellType() != XSSFCell.CELL_TYPE_STRING) {
                    continue;
                }

                String cellstr = cell.getStringCellValue();
                if (null == cellstr || "".equals(cellstr)) {
                    continue;
                }

                if (cellstr.startsWith(VALUED_COLUMN)) {
                    if (isFirst) {
                        rowAndColumnAndSize[0] = i;
                        rowAndColumnAndSize[1] = j;
                        isFirst = false;
                    }
                    ++size;
                }
            }
        }

        return rowAndColumnAndSize;
    }

    public static Integer findLoopRowNumberNoMerged(Sheet sheet, int firstRow,
            int lastRow) {

        for (int i = firstRow; i <= lastRow; i++) {
            Row row = sheet.getRow(i);
            short firstCellNum = row.getFirstCellNum();
            short lastCellNum = row.getLastCellNum();

            for (short j = firstCellNum; j <= lastCellNum; j++) {
                Cell cell = row.getCell(j);

                if (null == cell
                        || cell.getCellType() != XSSFCell.CELL_TYPE_STRING) {
                    continue;
                }

                String cellstr = cell.getStringCellValue();
                if (null == cellstr || "".equals(cellstr)) {
                    continue;
                }

                if (cellstr.startsWith(VALUED_DELIM)
                        && cellstr.endsWith(VALUED_DELIM2)) {
                    if (WorkbookUtils.getCellMergedIndex(sheet, i,
                            cell.getColumnIndex()) == 0) {
                        return i;
                    }
                }
            }
        }

        return 0;
    }

    /**
     * 拷贝多行
     */
    public static void copyMutilRow(Sheet sheet, int from, int count,
            int twoCount) {
        int cellRangeIndex = cellMergedForLoop(sheet, from);
        switch (cellRangeIndex) {
        case -1:
            break;
        case 0:
            WorkbookUtils.copyRow(sheet, from, from + 1, count - 1);
            break;
        default:
            CellRangeAddress cell = sheet.getMergedRegion(cellRangeIndex);
            int size = cell.getLastRow() - cell.getFirstRow() + 1;
            int rowNum = findLoopRowNumber(sheet, from + 1, from + size);

            switch (rowNum) {
            case 0:
                WorkbookUtils.copyCellRangeRow(sheet, cellRangeIndex, count, 0);
                break;
            default:
                if (twoCount > 1) { // 大于1条才拷贝
                    WorkbookUtils.copyRow(sheet, rowNum, rowNum + 1,
                            twoCount - 1);
                }

                WorkbookUtils.copyCellRangeRow(sheet, cellRangeIndex, count,
                        twoCount);
            }
        }

    }

    /**
     * 拷贝列
     */
    public static void copyColumn(Sheet sheet, int rowNum, int lastRow, int columnNum,
            int copySize, List<Map<String, String>> data) {

        if (data != null && !data.isEmpty()) {
            int dataSize = data.size() - 1;
            if (dataSize > 0) {
                for(int r = rowNum; r <= lastRow; r ++) {
                    Row row = sheet.getRow(r);
                    if (null == row) {
                        return;
                    }
                    for (int i = 1; i <= dataSize; i++) {
                        for (int j = 0; j < copySize; j++) {
                            Cell fromCell = row.getCell(columnNum + j);
                            if (null != fromCell) {
                                Cell toCell = WorkbookUtils.getCell(row, columnNum + copySize * i
                                        + j);
                                toCell.setCellStyle(fromCell.getCellStyle());
                                toCell.setCellType(fromCell.getCellType());
                                sheet.autoSizeColumn(columnNum + copySize * i+ j, true);
                                sheet.setColumnWidth(columnNum + copySize * i+ j, 8888);
                                toCell.setCellValue(fromCell.getStringCellValue());
                            }

                        }
                    }
                    
                }

            }

        } else {
            for(int r = rowNum; r <= lastRow; r ++) {
                Row row = sheet.getRow(r);
                for (int i = 0; i < copySize; i++) {
                    Cell cell = row.getCell(columnNum + i);
                    if (null != cell) {
                        row.removeCell(cell);
                    }

                }
            }
        }

    }
    
    /**
     * 设置列值
     */
    public static void setColumnLoopData(Row row, int columnNum, int copySize, Map<String, String> oneData) {
        Set<String> keys = oneData.keySet() ;
        for (int i = 0; i < copySize; i++) {
            Cell cell = row.getCell(columnNum + i);
            if (null == cell
                    || cell.getCellType() != XSSFCell.CELL_TYPE_STRING) {
                continue;
            }

            String cellstr = cell.getStringCellValue();
            if (null == cellstr || "".equals(cellstr)) {
                continue;
            }

            if (cellstr.contains(VALUED_COLUMN)
                    && cellstr.contains(VALUED_DELIM2)) {
                int startIndex = cellstr.indexOf(VALUED_COLUMN)
                        + VALUED_COLUMN.length();
                int endIndex = cellstr.indexOf(VALUED_DELIM2)
                        + VALUED_DELIM2.length();
                String start = cellstr.substring(0,
                        startIndex - VALUED_COLUMN.length());
                String value = cellstr.substring(startIndex, endIndex
                        - VALUED_DELIM2.length());
                String end = cellstr.substring(endIndex);

                if (keys.contains(value)) {
                    cell.setCellValue(new StringBuilder(start)
                            .append(oneData.get(value)).append(end).toString());
                }
            }

        } 
    }

    /**
     * 判断拷贝行开头是否是合并单元格, 返回单元格的总行数, 0代表不是拷贝行
     */
    public static int cellMergedForLoop(Sheet sheet, int from) {
        Row fromRow = sheet.getRow(from);
        Short firstCellNum = fromRow.getFirstCellNum();
        Short lastCellNum = fromRow.getLastCellNum();
        for (int i = firstCellNum; i < lastCellNum; i++) {
            Cell cell = fromRow.getCell(i);
            if (null == cell || cell.getCellType() != XSSFCell.CELL_TYPE_STRING) {
                continue;
            }

            String cellstr = cell.getStringCellValue();
            if (null == cellstr || "".equals(cellstr)) {
                continue;
            }

            if (cellstr.startsWith(VALUED_DELIM)
                    && cellstr.endsWith(VALUED_DELIM2)) {
                int index = WorkbookUtils.getCellMergedIndex(sheet, from, i);
                if (index == 0) {
                    return 0;
                } else {
                    return index;
                }
            }
        }

        return -1;
    }

}
