package com.neunn.excelutils.excelentity;

import java.util.HashMap;
import java.util.Map;

/**
 * @author hjl
 * @E-mail:huangjl@neunn.com
 * @version 创建时间：2015年10月19日 下午1:13:15
 */
public class ExcelExportEntity {
    
    private Map<String, ExcelExportSheetEntity> excelData = new HashMap<String, ExcelExportSheetEntity>() ;

    public Map<String, ExcelExportSheetEntity> getExcelData() {
        return excelData;
    }

    public void setExcelData(Map<String, ExcelExportSheetEntity> excelData) {
        this.excelData = excelData;
    } 
}
