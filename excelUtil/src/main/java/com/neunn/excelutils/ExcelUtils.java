package com.neunn.excelutils;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;

import javax.servlet.ServletContext;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.neunn.excelutils.excelentity.ExcelExportEntity;
import com.neunn.excelutils.excelentity.ExcelExportSheetEntity;

/**
 * @author hjl
 * @E-mail:huangjl@neunn.com
 * @version 创建时间：2015年10月19日 下午5:51:57
 */
public class ExcelUtils {
    
    public static void exportWb(ServletContext ctx, String config,
            ExcelExportEntity context, OutputStream out) throws ExcelException {
        try {
            Workbook wb = WorkbookUtils.openWorkbook(ctx, config);
            parseWorkbook(context, wb);
            wb.write(out);
        } catch (Exception e) {
            throw new ExcelException(e.getMessage());
        }
        
    }
    
    public static void exportWb(String fileName, ExcelExportEntity context,
            OutputStream out) throws ExcelException {
        try {
            Workbook wb = WorkbookUtils.openWorkbook(fileName);
            parseWorkbook(context, wb);
            wb.write(out);
        } catch (Exception e) {
            throw new ExcelException(e.getMessage());
        }
    }

    public static void exportWb(InputStream inputStream, ExcelExportEntity context,
            OutputStream out) throws ExcelException {
        try {
            Workbook wb = WorkbookUtils.openWorkbook(inputStream);
            parseWorkbook(context, wb);
            wb.write(out);
        } catch (Exception e) {
            throw new ExcelException(e.getMessage());
        }
    }
    
    public static void exportWb(OPCPackage pkg, ExcelExportEntity context,
            OutputStream out) throws ExcelException {
        try {
            Workbook wb = WorkbookUtils.openWorkbook(pkg);
            parseWorkbook(context, wb);
            wb.write(out);
        } catch (Exception e) {
            throw new ExcelException(e.getMessage());
        }
    }
    
    public static void exportWb(File file, ExcelExportEntity context,
            OutputStream out) throws ExcelException {
        try {
            Workbook wb = WorkbookUtils.openWorkbook(file);
            parseWorkbook(context, wb);
            wb.write(out);
        } catch (Exception e) {
            throw new ExcelException(e.getMessage());
        }
    }
    public static void parseWorkbook(ExcelExportEntity context, Workbook wb)
            throws ExcelException {
        try {
            int sheetCount = wb.getNumberOfSheets();
            for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
                Sheet sheet = wb.getSheetAt(sheetIndex);
                ExcelExportSheetEntity entity = context.getExcelData().get(sheetIndex + "") ;
                parseSheet(entity, sheet);
            }
        } catch (Exception e) {
            throw new ExcelException(e.getMessage());
        }
    }

    public static void parseSheet(ExcelExportSheetEntity entity, Sheet sheet)
            throws ExcelException {
        try {
            ExcelParser.parse(entity, sheet);
        } catch (Exception e) {
            e.printStackTrace();
            throw new ExcelException(e.getMessage());
        } 
    }

    public static boolean isCanShowType(Object value) {
        if (null == value)
            return false;
        String valueType = value.getClass().getName();
        return "java.lang.String".equals(valueType)
                || "java.lang.Double".equals(valueType)
                || "java.lang.Integer".equals(valueType)
                || "java.lang.Boolean".equals(valueType)
                || "java.sql.Timestamp".equals(valueType)
                || "java.util.Date".equals(valueType)
                || "java.lang.Byte".equals(valueType)
                || "java.math.BigDecimal".equals(valueType)
                || "java.math.BigInteger".equals(valueType)
                || "java.lang.Float".equals(valueType)
                || value.getClass().isPrimitive();
    }

}
