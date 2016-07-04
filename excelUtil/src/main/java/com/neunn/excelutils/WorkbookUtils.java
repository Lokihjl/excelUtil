package com.neunn.excelutils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.servlet.ServletContext;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author hjl
 * @E-mail:huangjl@neunn.com
 * @version 创建时间：2015年10月19日 下午5:37:32
 */
public class WorkbookUtils {

    public WorkbookUtils() {
    }

    /**
     * 通过ServletContext打开wb
     * 
     */
    public static Workbook openWorkbook(ServletContext ctx,
            String config) throws ExcelException {
        InputStream in = null;
        Workbook wb = null;
        try {
            in = ctx.getResourceAsStream(config);
            wb = new XSSFWorkbook(in);
        } catch (Exception e) {
            throw new ExcelException("File" + config + "not found,"
                    + e.getMessage());
        } finally {
            try {
                in.close();
            } catch (Exception e) {

            }
        }

        return wb;
    }

    /**
     * 通过路径获取 wb
     * 
     */
    public static Workbook openWorkbook(String fileName)
            throws ExcelException {
        InputStream in = null;
        Workbook wb = null;
        try {
            in = new FileInputStream(fileName);
            wb = new XSSFWorkbook(in);
        } catch (Exception e) {
            throw new ExcelException("File" + fileName + "not found"
                    + e.getMessage());
        } finally {
            try {
                in.close();
            } catch (Exception e) {

            }
        }

        return wb;
    }
    
    public static Workbook openWorkbook(OPCPackage pkg)
            throws ExcelException {
        Workbook wb = null;
        try {
            wb = new XSSFWorkbook(pkg);
        } catch (Exception e) {
            throw new ExcelException(e.getMessage());
        }
        return wb;
    }

    /**
     * 通过流开启wb
     * 
     */
    public static Workbook openWorkbook(InputStream in)
            throws ExcelException {
        Workbook wb = null;
        try {
            wb = new XSSFWorkbook(in);
        } catch (Exception e) {
            throw new ExcelException(e.getMessage());
        }
        return wb;
    }
    
    public static Workbook openWorkbook(File file) throws ExcelException {
        Workbook wb = null;
        try {
            wb = new XSSFWorkbook(file);
        } catch (Exception e) {
            throw new ExcelException(e.getMessage());
        }
        return wb;
    }

    /**
     * 保存Excel到流中
     */
    public static void SaveWorkbook(Workbook wb, OutputStream out)
            throws ExcelException {
        try {
            wb.write(out);
        } catch (Exception e) {
            throw new ExcelException(e.getMessage());
        }
    }

    /**
     * 设置单元格内容
     */
    public static void setCellValue(Sheet sheet, int rowNum, int colNum,
            String value) {
        Row row = getRow(rowNum, sheet);
        Cell cell = getCell(row, colNum);
        cell.setCellValue(value);
    }

    /**
     * 获取单元格内容
     */
    public static String getStringCellValue(Sheet sheet, int rowNum,
            int colNum) {
        Row row = getRow(rowNum, sheet);
        Cell cell = getCell(row, colNum);
        return cell.getStringCellValue();
    }

    /**
     * 设置double类型数据到单元格中
     */
    public static void setCellValue(Sheet sheet, int rowNum, int colNum,
            double value) {
        Row row = getRow(rowNum, sheet);
        Cell cell = getCell(row, colNum);
        cell.setCellValue(value);
    }

    /**
     * 获取单元格内容
     */
    public static double getNumericCellValue(Sheet sheet, int rowNum,
            int colNum) {
        Row row = getRow(rowNum, sheet);
        Cell cell = getCell(row, colNum);
        return cell.getNumericCellValue();
    }

    /**
     * 设置日期类型的数据到单元格中
     */
    public static void setCellValue(Sheet sheet, int rowNum, int colNum,
            Date value) {
        Row row = getRow(rowNum, sheet);
        Cell cell = getCell(row, colNum);
        cell.setCellValue(value);
    }

    /**
     * 获取单元格日期类型的数据
     */
    public static Date getDateCellValue(Sheet sheet, int rowNum, int colNum) {
        Row row = getRow(rowNum, sheet);
        Cell cell = getCell(row, colNum);
        return cell.getDateCellValue();
    }

    /**
     * 设置Boolean类型的数据到单元格中
     */
    public static void setCellValue(Sheet sheet, int rowNum, int colNum,
            boolean value) {
        Row row = getRow(rowNum, sheet);
        Cell cell = getCell(row, colNum);
        cell.setCellValue(value);
    }

    /**
     * 获取单元格中Boolean类型的数据
     * 
     */
    public static boolean getBooleanCellValue(Sheet sheet, int rowNum,
            int colNum) {
        Row row = getRow(rowNum, sheet);
        Cell cell = getCell(row, colNum);
        return cell.getBooleanCellValue();
    }

    /**
     * 获取sheet中的行，如果不存在则创建
     * 
     */
    public static Row getRow(int rowCounter, Sheet sheet) {
        Row row = sheet.getRow((short) rowCounter);
        if (row == null) {
            row = sheet.createRow((short) rowCounter);
        }
        return row;
    }

    /**
     * 获取单元格如果不存在则创建
     */
    public static Cell getCell(Row row, int column) {
        Cell cell = row.getCell((short) column);

        if (cell == null) {
            cell = row.createCell((short) column);
        }
        return cell;
    }

    /**
     * 获取单元格如果不存在则创建
     */
    public static Cell getCell(Sheet sheet, int rowNum, int colNum) {
        Row row = getRow(rowNum, sheet);
        Cell cell = getCell(row, colNum);
        return cell;
    }
    
    /**
     * 拷贝以合并单元开头的行
     */
    public static void copyCellRangeRow(Sheet sheet, int cellRangeIndex, int count, int twoCount) {
        CellRangeAddress cellRange = sheet.getMergedRegion(cellRangeIndex) ;
        int firstRow = cellRange.getFirstRow() ;
        int lastRow = cellRange.getLastRow() ;
        int firstColumn = cellRange.getFirstColumn() ;
        int lastColumn = cellRange.getLastColumn() ;
        twoCount = twoCount == 0 ? 1 : twoCount ;
        int size = lastRow - firstRow + twoCount;
        
        // 去掉单元格
        sheet.removeMergedRegion(cellRangeIndex);
        
        // 拷贝行
        for(int i = 1 ; i < count; i ++) {
            for(int j = 0; j < size; j ++) {
                copyRow(sheet, firstRow + j, firstRow + (size * i) + j, 1);
            }
            CellRangeAddress newCellRange = new CellRangeAddress(firstRow + size * i, firstRow + size * (i + 1) - 1, firstColumn, lastColumn) ;
            sheet.addMergedRegion(newCellRange) ;
        }
        
        cellRange = new CellRangeAddress(firstRow, lastRow + twoCount - 1, firstColumn, lastColumn) ;
        // 合并单元格
        sheet.addMergedRegion(cellRange) ;
    }
    
    /**
     * 行拷贝
     */
    public static void copyRow(Sheet sheet, int from, int to, int count) {

        for (int rownum = from; rownum < from + count; rownum++) {
            Row fromRow = sheet.getRow(rownum);
            Row toRow = getRow(to + rownum - from, sheet);
            if (null == fromRow)
                return;
            toRow.setHeight(fromRow.getHeight());
            toRow.setHeightInPoints(fromRow.getHeightInPoints());
            short lastCellNum = fromRow.getLastCellNum() ;
            for (int i = fromRow.getFirstCellNum(); i <= lastCellNum && i >= 0; i++) {
                Cell fromCell = getCell(fromRow, i);
                Cell toCell = getCell(toRow, i);
                toCell.setCellStyle(fromCell.getCellStyle());
                toCell.setCellType(fromCell.getCellType());
                switch (fromCell.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    toCell.setCellValue(fromCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    toCell.setCellFormula(fromCell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    toCell.setCellValue(fromCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    toCell.setCellValue(fromCell.getStringCellValue());
                    break;
                default:
                }
            }
        }

        // copy merged region
        List<CellRangeAddress> shiftedRegions = new ArrayList<CellRangeAddress>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress r = sheet.getMergedRegion(i);

            if (r.getFirstRow() >= from && r.getLastRow() < from + count) {
                CellRangeAddress n_r = new CellRangeAddress(r.getFirstRow()
                        + to - from, r.getLastRow() + to - from,
                        r.getFirstColumn(), r.getLastColumn());
                shiftedRegions.add(n_r);
            }
        }

        // readd so it doesn't get shifted again
        for (CellRangeAddress cellRange : shiftedRegions) {
            sheet.addMergedRegion(cellRange);
        }
    }
    
    // 移动单元格
    public static void shiftCell(Sheet sheet, Row row,
            Cell beginCell, int shift, int rowCount) {

        if (shift == 0)
            return;

        // get the from & to row
        int fromRow = row.getRowNum();
        int toRow = row.getRowNum() + rowCount - 1;
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress r = sheet.getMergedRegion(i);
            if (r.getFirstRow() == row.getRowNum()) {
                if (r.getLastRow() > toRow) {
                    toRow = r.getLastRow();
                }
                if (r.getFirstRow() < fromRow) {
                    fromRow = r.getFirstRow();
                }
            }
        }

        for (int rownum = fromRow; rownum <= toRow; rownum++) {
            Row curRow = WorkbookUtils.getRow(rownum, sheet);
            int lastCellNum = curRow.getLastCellNum();
            for (int cellpos = lastCellNum; cellpos >= beginCell
                    .getColumnIndex(); cellpos--) {
                Cell fromCell = WorkbookUtils.getCell(curRow, cellpos);
                Cell toCell = WorkbookUtils
                        .getCell(curRow, cellpos + shift);
                toCell.setCellType(fromCell.getCellType());
                toCell.setCellStyle(fromCell.getCellStyle());
                switch (fromCell.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    toCell.setCellValue(fromCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    toCell.setCellFormula(fromCell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    toCell.setCellValue(fromCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    toCell.setCellValue(fromCell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR:
                    toCell.setCellErrorValue(fromCell.getErrorCellValue());
                    break;
                }
                fromCell.setCellValue("");
                fromCell.setCellType(Cell.CELL_TYPE_BLANK);
                Workbook wb = new XSSFWorkbook();
                CellStyle style = wb.createCellStyle();
                fromCell.setCellStyle(style);
            }

            // process merged region
            for (int cellpos = lastCellNum; cellpos >= beginCell
                    .getColumnIndex(); cellpos--) {
                Cell fromCell = WorkbookUtils.getCell(curRow, cellpos);

                List<CellRangeAddress> shiftedRegions = new ArrayList<CellRangeAddress>();
                for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                    CellRangeAddress r = sheet.getMergedRegion(i);
                    if (r.getFirstRow() == curRow.getRowNum()
                            && r.getFirstColumn() == fromCell.getColumnIndex()) {

                        r.setFirstColumn(r.getFirstColumn() + shift);
                        r.setLastColumn(r.getLastColumn() + shift);

                        // have to remove/add it back
                        shiftedRegions.add(r);
                        sheet.removeMergedRegion(i);
                        // we have to back up now since we removed one
                        i = i - 1;
                    }
                }

                // readd so it doesn't get shifted again
                for (CellRangeAddress cellRange : shiftedRegions) {
                    sheet.addMergedRegion(cellRange);
                }
            }
        }
    }
    
    /**
     * 获取合并单元格 0 为不是合并单元格
     */
    public static int getCellMergedIndex(Sheet sheet, int startRowNum, int startColumn) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress cell = sheet.getMergedRegion(i);

            if(cell.getFirstRow() == startRowNum && cell.getFirstColumn() == startColumn) {
                return i ;
            }
        }
        return 0;
    }

    /**
     * 删除行
     * 
     */
    public void removeRow(Sheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        }
        if (rowIndex == lastRowNum) {
            Row removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }

    /**
     * 添加图片
     */
    public static void addPicture(Workbook wb, Sheet sheet, String picFileName,
            int picType, int row, int col) {
        InputStream is = null;
        try {
            // 读取图片
            is = new FileInputStream(picFileName);
            byte[] bytes = IOUtils.toByteArray(is);
            int pictureIdx = wb.addPicture(bytes, picType);
            is.close();
            // 写图片
            CreationHelper helper = wb.getCreationHelper();
            Drawing drawing = sheet.createDrawingPatriarch();
            ClientAnchor anchor = helper.createClientAnchor();
            // 设置图片的位置
            anchor.setCol1(col);
            anchor.setRow1(row);
            Picture pict = drawing.createPicture(anchor, pictureIdx);

            pict.resize();
        } catch (Exception e) {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (IOException e1) {
                e1.printStackTrace();
            }
            e.printStackTrace();
        }
    }

    /**
     * 创建Cell 默认为水平和垂直方式都是居中
     */
    public static Cell createCell(CellStyle style, Row row, short column) {
        return createCell(style, row, column, CellStyle.ALIGN_CENTER,
                CellStyle.ALIGN_CENTER);
    }

    /**
     * 创建Cell并设置水平和垂直方式
     */
    public static Cell createCell(CellStyle style, Row row, short column,
            short halign, short valign) {
        Cell cell = row.createCell(column);
        setAlign(style, halign, valign);
        cell.setCellStyle(style);
        return cell;
    }

    /**
     * 设置单元格对齐方式
     */
    public static CellStyle setAlign(CellStyle style, short halign, short valign) {
        style.setAlignment(halign);
        style.setVerticalAlignment(valign);
        return style;
    }

    /**
     * 设置单元格边框(四个方向的颜色一样)
     */
    public static CellStyle setBorder(CellStyle style, short borderStyle,
            short borderColor) {

        // 设置底部格式（样式+颜色）
        style.setBorderBottom(borderStyle);
        style.setBottomBorderColor(borderColor);
        // 设置左边格式
        style.setBorderLeft(borderStyle);
        style.setLeftBorderColor(borderColor);
        // 设置右边格式
        style.setBorderRight(borderStyle);
        style.setRightBorderColor(borderColor);
        // 设置顶部格式
        style.setBorderTop(borderStyle);
        style.setTopBorderColor(borderColor);

        return style;
    }

    /**
     * 设置前景颜色
     */
    public static CellStyle setBackColor(CellStyle style, short color) {

        // 设置前端颜色
        style.setFillForegroundColor(color);
        // 设置填充模式
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);

        return style;
    }

    /**
     * 设置背景颜色
     */
    public static CellStyle setBackColor(CellStyle style, short backColor,
            short fillPattern) {

        // 设置背景颜色
        style.setFillBackgroundColor(backColor);

        // 设置填充模式
        style.setFillPattern(fillPattern);

        return style;
    }

    /**
     * 设置字体
     */
    public static CellStyle setFont(Font font, CellStyle style, short fontSize,
            short color, String fontName) {
        font.setFontHeightInPoints(color);
        font.setFontName(fontName);

        // font.setItalic(true);// 斜体
        // font.setStrikeout(true);//加干扰线

        font.setColor(color);// 设置颜色
        // Fonts are set into a style so create a new one to use.
        style.setFont(font);

        return style;

    }
    
//    private static void replaceHeader(Sheet sheet, Map<String, String> replaceTextMap) {
//        // 遍历所有行
//        for (Row thisRow : sheet) {
//            boolean isFound = false;
//            // 便利所有列
//            float defaultRowHeight = thisRow.getHeightInPoints();
//            float maxHeight = defaultRowHeight;
//            for (Cell thisCell : thisRow) {
//        
//                
//                // 获取单元格的类型
//                CellReference cellRef = new CellReference(thisRow.getRowNum(),
//                        thisCell.getColumnIndex());
//                switch (thisCell.getCellType()) {
//                // 字符串
//                case Cell.CELL_TYPE_STRING:
//                    String targetText = thisCell.getRichStringCellValue()
//                            .getString();
//                    
//
//                    
//                    if (targetText != null && !targetText.trim().equals("")) {
//                        if(replaceTextMap.containsKey(targetText)){
//                            
//                        float thisHeight = getExcelCellAutoHeight((String)replaceTextMap.get(targetText),defaultRowHeight, getMergedCellNum(thisCell) * sheet.getColumnWidth(thisCell.getColumnIndex())/256); 
//                        if(thisHeight > maxHeight) maxHeight = thisHeight;
//                            isFound = true;
//                            thisCell.setCellValue((String)replaceTextMap.get(targetText));
//                            thisCell.getCellStyle().setWrapText(true);
//                        }
//                    }
//
//                    break;
//                // 数字
//                case Cell.CELL_TYPE_NUMERIC:
//                    break;
//                // boolean
//                case Cell.CELL_TYPE_BOOLEAN:
//                    break;
//                // 方程式
//                case Cell.CELL_TYPE_FORMULA:
//                    break;
//                // 空值
//                default:
//                }
//
//            }
//            
//            if(isFound) thisRow.setHeightInPoints(maxHeight);
//        }
//
//    }

}
