package util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.Date;

public final class XLSUtil {
    /**
     *
     * @param worksheet
     * @param line
     * @param col
     * @param value
     */
    public static void setCellXY(HSSFSheet worksheet, int line, int col, LocalDate value) {
        Cell cell = worksheet.getRow( (line - 1) ).getCell( (col - 1) );   // Access the second cell in second row to update the value
        Date date = java.sql.Date.valueOf(value);
        cell.setCellValue(date);  // Get current cell value value and overwrite the value
    }

    /**
     *
     * @param worksheet
     * @param line
     * @param col
     * @param value
     */
    public static void setCellXY(HSSFSheet worksheet, int line, int col, Integer value) {
        double val = value.doubleValue();
        setCellXY(worksheet, line, col, val);
    }

    /**
     *
     * @param worksheet
     * @param line
     * @param col
     * @param value
     */
    public static void setCellXY(HSSFSheet worksheet, int line, int col, BigDecimal value) {
        double val = value.doubleValue();
        setCellXY(worksheet, line, col, val);
    }

    /**
     *
     * @param worksheet
     * @param line
     * @param col
     * @param value
     */
    public static void setCellXY(HSSFSheet worksheet, int line, int col, double value) {
        Cell cell = worksheet.getRow( (line - 1) ).getCell( (col - 1) );   // Access the second cell in second row to update the value
        cell.setCellValue(value);  // Get current cell value value and overwrite the value
    }

    /**
     *
     * @param worksheet
     * @param line
     * @param col
     * @param value
     */
    public static void setCellXY(HSSFSheet worksheet, int line, int col, String value) {
        Cell cell = worksheet.getRow( (line - 1) ).getCell( (col - 1) );   // Access the second cell in second row to update the value
        cell.setCellValue(value);  // Get current cell value value and overwrite the value
    }

    /**
     * Este método foi copiado do site
     * http://www.zachhunter.com/2010/05/npoi-copy-row-helper/
     *
     * @param workbook
     * @param worksheet
     * @param sourceRowNum
     * @param destinationRowNum
     */
    public static void copyRow(HSSFWorkbook workbook, HSSFSheet worksheet, int sourceRowNum, int destinationRowNum) {
        // Get the source / new row
        HSSFRow newRow = worksheet.getRow(destinationRowNum);
        HSSFRow sourceRow = worksheet.getRow(sourceRowNum);

        // If the row exist in destination, push down all rows by 1 else create a new row
        if (newRow != null) {
            worksheet.shiftRows(destinationRowNum, worksheet.getLastRowNum(), 1);
        } else {
            newRow = worksheet.createRow(destinationRowNum);
        }

        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            HSSFCell oldCell = sourceRow.getCell(i);
            HSSFCell newCell = newRow.createCell(i);

            // If the old cell is null jump to next cell
            if (oldCell == null) {
                newCell = null;
                continue;
            }

            // Copy style from old cell and apply to new cell
            HSSFCellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            ;
            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());

            // Set the cell data value
            switch (oldCell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
            }
        }

        // If there are are any merged regions in the source row, copy to new row
        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                    (newRow.getRowNum() +
                        (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()
                        )),
                    cellRangeAddress.getFirstColumn(),
                    cellRangeAddress.getLastColumn());
                worksheet.addMergedRegion(newCellRangeAddress);
            }
        }
    }
}