package com.xinbang.check.core.util;

import com.google.common.collect.Lists;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;

import java.io.*;
import java.util.*;

/**
 * 分割excel文件
 * Created by Yitar on 2016/9/28.
 */
@Slf4j
public class ExcelSplitUtil {

    public static void main(String[] args) {
        splitExcel("E:/loan_2016-10-01.xls", 1);
    }

    public static List<String> splitExcel(String fileName, int grained) {

        log.info("开始分割文件");

        try {
            List<String> urlList;
            List<HSSFWorkbook> workbookList = getSplitMap(fileName, grained);
            urlList = createSplitXSSFWorkbook(workbookList, fileName);
            return urlList;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }


    private static List<HSSFWorkbook> getSplitMap(String fileName, int grained) throws Exception {

        List<HSSFWorkbook> workBookList = Lists.newArrayList();
        InputStream is = new FileInputStream(new File(fileName));
        HSSFWorkbook workBook = new HSSFWorkbook(is);
        HSSFSheet sheet = workBook.getSheetAt(0);
        int rowNum = sheet.getPhysicalNumberOfRows() - 1;
        HSSFRow titleRow = sheet.getRow(0);
        int column = titleRow.getLastCellNum();

        int numberOfFiles = (rowNum / grained) * grained >= rowNum ? rowNum / grained : rowNum / grained + 1;


        for (int i = 0; i < numberOfFiles; i++) {
            HSSFWorkbook newWorkBook = new HSSFWorkbook();
            HSSFSheet newSheet = newWorkBook.createSheet();
            HSSFRow fistRow = newSheet.createRow(0);
            titleRow.forEach(x -> fistRow.createCell(x.getColumnIndex()).setCellValue(x.getStringCellValue()));
            workBookList.add(newWorkBook);
        }

        if (rowNum > grained) {

            for (int i = 0; i < numberOfFiles; i++) {

                HSSFWorkbook tempWorkBook = workBookList.get(i);
                HSSFSheet tempSheet = tempWorkBook.getSheetAt(0);

                int start = getStartIndex(numberOfFiles, grained, i) + 1;

                log.info("{}", getStartIndex(numberOfFiles, grained, i + 1) + 1);

                int end = getStartIndex(numberOfFiles, grained, i + 1) + 1
                        >= rowNum ? rowNum
                        : getStartIndex(numberOfFiles, grained, i + 1) + 1;

                int count = 1;
                if (end != start && count <= end - start || end == start) {

                    if (end - start < grained) {

                        for (int j = start; j < end + 1; j++) {
                            HSSFRow tempRow = tempSheet.createRow(count++);

                            for (int k = 0; k < column + 1; k++) {
                                setCellValue(tempRow.createCell(k), sheet.getRow(j).getCell(k), workBook);
                            }
                        }

                    } else {
                        for (int j = start; j < end; j++) {

                            HSSFRow tempRow = tempSheet.createRow(count++);

                            for (int k = 0; k < column + 1; k++) {
                                setCellValue(tempRow.createCell(k), sheet.getRow(j).getCell(k), workBook);
                            }


                        }
                    }

                }
                workBookList.set(i, tempWorkBook);
            }
            return workBookList;
        }
        return workBookList;
    }

    private static List<String> createSplitXSSFWorkbook(List<HSSFWorkbook> workbookList, String fileName)
            throws IOException {

        List<String> urlList = Lists.newArrayList();
        String fileToSavePath = fileName.substring(0, fileName.lastIndexOf("/"));
        File filePath = new File(fileToSavePath);

        if (!filePath.exists()) {
            Boolean flag = filePath.mkdir();
            if (!flag) {
                return null;
            }
        }

        if (!filePath.isDirectory()) {
            System.out.println("无效文件目录");
            return null;
        }

        if (workbookList.size() == 1) {
            System.out.println("文件无需分割");
            urlList.add(fileName);
        } else {

            int i = 0;

            for (Workbook workbook : workbookList) {

                try {
                    FileOutputStream fOut;

                    String newFilePath = filePath + "/" +
                            com.xinbang.utils.DateUtil.getCurrentDate(com.xinbang.utils.DateUtil.fullPatterns)
                                    .toString().replaceAll(" ", "").replaceAll(":", "") + "_" + i++ + ".xls";

                    File file = new File(newFilePath);

                    fOut = new FileOutputStream(file);
                    workbook.write(fOut);
                    fOut.flush();
                    fOut.close();

                    urlList.add(newFilePath);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return urlList;
    }

    private static void setCellValue(Cell newCell, Cell cell, HSSFWorkbook wb) {
        if (cell == null) {
            return;
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                newCell.setCellValue(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    HSSFCellStyle cellStyle = wb.createCellStyle();
                    HSSFDataFormat format = wb.createDataFormat();
                    cellStyle.setDataFormat(format.getFormat("yyyy/mm/dd HH:mm:ss"));
                    newCell.setCellStyle(cellStyle);
                    newCell.setCellValue(cell.getDateCellValue());
                } else {
                    newCell.setCellValue(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_FORMULA:
                newCell.setCellValue(cell.getCellFormula());
                break;
            case Cell.CELL_TYPE_STRING:
                newCell.setCellValue(cell.getStringCellValue());
                break;
        }

    }

    private static int getStartIndex(int numberOfFiles, int grained, int i) {

        if (i > numberOfFiles) {
            i = numberOfFiles;
        }
        return i * grained;
    }
}
