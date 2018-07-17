package util;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.util.Date;

/**
 * @author Jonatan Raimir
 */
public final class XLSUtil {
    private String pathFileModel;
    private FileInputStream fileInputStream;
    private HSSFWorkbook workbook;
    private HSSFSheet worksheet;


    public void openFileXLS(File file, int sheetAt) throws IOException {
        this.fileInputStream = new FileInputStream(file);
        this.workbook = new HSSFWorkbook(fileInputStream); //Access the workbook
        this.worksheet = workbook.getSheetAt(sheetAt); //Access the worksheet, so that we can update / modify it.
    }

    public void openFileXLS(File file, String nameSheet) throws IOException {
        this.fileInputStream = new FileInputStream(file);
        this.workbook = new HSSFWorkbook(fileInputStream); //Access the workbook
        this.worksheet = workbook.getSheet(nameSheet); //Access the worksheet, so that we can update / modify it.
    }

    public void openFileXLS(String pathFile, int sheetAt) throws IOException {
        __openFileXLS(pathFile);
        this.worksheet = workbook.getSheetAt(sheetAt); //Access the worksheet, so that we can update / modify it.
    }

    public void openFileXLS(String pathFile, String nameSheet) throws IOException {
        __openFileXLS(pathFile);
        this.worksheet = workbook.getSheet(nameSheet); //Access the worksheet, so that we can update / modify it.
    }

    private void __openFileXLS(String pathFile) throws IOException {
        this.pathFileModel = pathFile;
        this.fileInputStream = new FileInputStream(new File(pathFile));
        this.workbook = new HSSFWorkbook(fileInputStream); //Access the workbook
    }

    public void closeFile() throws IOException {
        getFileInputStream().close();
    }

    public static String createTemporaryCopyFile(HSSFWorkbook workbook) throws IOException {
        Timestamp timestamp = new Timestamp(System.currentTimeMillis());
        String pathNewTemporaryFile =  FileUtils.getTempDirectoryPath() + "newFileTemporary_" + timestamp.getTime() + ".xls";
        FileOutputStream outputFile = new FileOutputStream(new File(pathNewTemporaryFile));  //Open FileOutputStream to write updates
        workbook.write(outputFile); //write changes
        outputFile.close();  //close the stream
        return pathNewTemporaryFile;
    }
    
    /**
     * deleteFile
     * @param pathFile
     */
    public static void deleteFile(String pathFile) {
       File file = new File(pathFile);

        if (file.delete()) {
            System.out.println(XLSUtil.class.getName() + " - File deleted successfully");
        }
        else {
            System.out.println(XLSUtil.class.getName() + " - Failed to delete the file");
        }
    }

    /**
     * Sets a value to a specific xls cell.
     * @param worksheet Chosen worksheet.
     * @param line The cell line.
     * @param col The cell column.
     * @param value The value to be bind.
     */
    public static void setCellXY(HSSFSheet worksheet, int line, int col, LocalDate value) {
        Cell cell = createCell(worksheet, line, col);
        if (value != null) {
            Date date = java.sql.Date.valueOf(value);
            cell.setCellValue(date);  // Get current cell value value and overwrite the value
        }
        else {
            cell.setCellValue("");
        }
    }

    /**
     * Sets a numeric value to a specific xls cell.
     * @param worksheet Chosen worksheet.
     * @param line The cell line.
     * @param col The cell column.
     * @param value The value to be bind.
     */
    public static void setCellXY(HSSFSheet worksheet, int line, int col, Number value) {
        if (value != null) {
            double val = value.doubleValue();
            setCellXY(worksheet, line, col, val);
        }
        else {
            setCellXY(worksheet, line, col, "");
        }
    }

    /**
     * Sets a value to a specific xls cell.
     * @param worksheet Chosen worksheet.
     * @param line The cell line.
     * @param col The cell column.
     * @param value The value to be bind.
     */
    public static void setCellXY(HSSFSheet worksheet, int line, int col, double value) {
        Cell cell = createCell(worksheet, line, col);
        cell.setCellValue(value);  // Get current cell value value and overwrite the value
    }

    /**
     * Sets a value to a specific xls cell.
     * @param worksheet Chosen worksheet.
     * @param line The cell line.
     * @param col The cell column.
     * @param value The value to be bind.
     */
    public static void setCellXY(HSSFSheet worksheet, int line, int col, String value) {
        Cell cell = createCell(worksheet, line, col);
        if (value != null) {
            cell.setCellValue(value);  // Get current cell value value and overwrite the value
        }
        else {
            cell.setCellValue("");
        }
    }

    /**
     * This method copies a line with its formatting.
     * Was taken from @link{http://www.zachhunter.com/2010/05/npoi-copy-row-helper/}.
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

    /**
     * create a cell in x,y coordinates
     * @param worksheet Chosen worksheet.
     * @param line The cell line.
     * @param col The cell column.
     */
    public static Cell createCell(HSSFSheet worksheet, int line, int col) {
        Row row = worksheet.getRow( --line );

        if (row == null) {
            row = worksheet.createRow( --line );
        }

        Cell cell = row.getCell( --col );

        //isNull
        if (cell == null) {
            cell = row.createCell(col);
        }

        return cell;
    }

    public String getPathFileModel() {
        return pathFileModel;
    }

    public HSSFWorkbook getWorkbook() {
        return workbook;
    }

    public HSSFSheet getWorksheet() {
        return worksheet;
    }

    public FileInputStream getFileInputStream() {
        return fileInputStream;
    }
}
