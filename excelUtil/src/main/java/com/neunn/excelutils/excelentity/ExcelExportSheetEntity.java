package com.neunn.excelutils.excelentity;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author hjl
 * @E-mail:huangjl@neunn.com
 * @version 创建时间：2015年10月19日 下午1:13:15
 */
public class ExcelExportSheetEntity {

    // 一次数据
    private Map<String, String> once = new HashMap<String, String>() ;
    
    // 循环数据
    private List<List<Map<String, String>>> manyRepeated ;
    
    private List<Map<String, String>> manyColumn ;
    
    // 合并单元格循环数据
    
    private Map<String, Map<String, String>> cellMergedOnce ;
    
    private Map<String, List<Map<String, String>>>  cellMergedRepeated ;
    
    public ExcelExportSheetEntity(Map<String, String> once, List<List<Map<String, String>>> manyRepeated,
            Map<String, Map<String, String>> cellMergedOnce,
            Map<String, List<Map<String, String>>>  cellMergedRepeated ,
            List<Map<String, String>> manyColumn) {
        this.once = once ;
        this.manyRepeated = manyRepeated ;
        this.cellMergedOnce = cellMergedOnce ;
        this.cellMergedRepeated = cellMergedRepeated ;
        this.manyColumn = manyColumn ;
    }

    public Map<String, String> getOnce() {
        return once;
    }

    public void setOnce(Map<String, String> once) {
        this.once = once;
    }

    public List<List<Map<String, String>>> getManyRepeated() {
        return manyRepeated;
    }

    public void setManyRepeated(List<List<Map<String, String>>> manyRepeated) {
        this.manyRepeated = manyRepeated;
    }

    public Map<String, Map<String, String>> getCellMergedOnce() {
        return cellMergedOnce;
    }

    public void setCellMergedOnce(Map<String, Map<String, String>> cellMergedOnce) {
        this.cellMergedOnce = cellMergedOnce;
    }

    public Map<String, List<Map<String, String>>> getCellMergedRepeated() {
        return cellMergedRepeated;
    }

    public void setCellMergedRepeated(
            Map<String, List<Map<String, String>>> cellMergedRepeated) {
        this.cellMergedRepeated = cellMergedRepeated;
    }

    public List<Map<String, String>> getManyColumn() {
        return manyColumn;
    }

    public void setManyColumn(List<Map<String, String>> manyColumn) {
        this.manyColumn = manyColumn;
    }
    
}
