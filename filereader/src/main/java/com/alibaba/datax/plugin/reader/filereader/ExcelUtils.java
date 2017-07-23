package com.alibaba.datax.plugin.reader.filereader;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author panyongjian
 * @since 2017/7/6.
 */
public class ExcelUtils {

    public static Workbook createWorkbook(InputStream inputstream){
        Workbook book = null;
        try {
            if (!(inputstream.markSupported())) {
                inputstream = new PushbackInputStream(inputstream, 8);
            }
            if (POIFSFileSystem.hasPOIFSHeader(inputstream)) {
                book = new HSSFWorkbook(inputstream);
            } else if (DocumentFactoryHelper.hasOOXMLHeader(inputstream)) {
                book = new XSSFWorkbook(OPCPackage.open(inputstream));
            }
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return book;
    }

    public static boolean isRowEmpty(Row row) {
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellTypeEnum() != CellType.BLANK)
                return false;
        }
        return true;
    }

    /**
     * 获取单元格内的值
     *
     * @param cell
     * @return
     */
    public static String toCellValueString(Cell cell) {
        if (cell == null) {
            return null;
        }
        String result = null;
        switch (cell.getCellTypeEnum()) {
            case STRING:
                result = cell.getRichStringCellValue() == null ? ""
                        : cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = formatDate(cell.getDateCellValue());
                } else {
                    result = readNumericCell(cell);
                }
                break;
            case BOOLEAN:
                result = Boolean.toString(cell.getBooleanCellValue());
                break;
            case BLANK:
                break;
            case ERROR:
                break;
            case FORMULA:
                try {
                    result = readNumericCell(cell);
                } catch (Exception e1) {
                    try {
                        result = cell.getRichStringCellValue() == null ? ""
                                : cell.getRichStringCellValue().getString();
                    } catch (Exception e2) {
                        throw new RuntimeException("获取公式类型的单元格失败", e2);
                    }
                }
                break;
            default:
                break;
        }
        return result;
    }

    private static String readNumericCell(Cell cell) {
        double value = cell.getNumericCellValue();
        return new BigDecimal(value).toString();
    }

    private static String formatDate(Date value) {
        if (value != null) {
            SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            return format.format(value);
        }
        return null;
    }

}
