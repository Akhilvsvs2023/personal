package apiautomation.globalgateway.generic;


import java.io.File;

import java.io.FileInputStream;

import java.io.IOException;

import java.io.InputStream;

import java.util.HashMap;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.ss.usermodel.WorkbookFactory;


public class ExcelReader {


    @SuppressWarnings("deprecation")

    public HashMap<String, String> readExcel(String sheetName, String testcase_name)

            throws IOException, InvalidFormatException {

        File file = new File("src/main/externals/InputData/InputData.xlsm");

        String filepath = file.getAbsolutePath();

        InputStream inp = new FileInputStream(filepath);

        Workbook wb = WorkbookFactory.create(inp);

        HashMap<String, String> input = new HashMap<String, String>();

        // get the sheet which needs read operation

        Sheet sh = wb.getSheet(sheetName);

        // get the total row count in the excel sheet

        int rowcount = sh.getLastRowNum();

        for (int i = 0; i <= rowcount; i++) {

            try {

                Row row = sh.getRow(i);

                // get the total cell count in the excel

                int cellcount = row.getLastCellNum();

                for (int j = 0; j < cellcount; j++) {

                    Cell cell = row.getCell(j);

                    cell.setCellType(Cell.CELL_TYPE_STRING);

                    // get cell value at the given position [i][j]

                    String value = cell.getStringCellValue();

                    // print the cell value

                    if (value.equals(testcase_name)) {

                        int cellcount_inner = sh.getRow(i + 1).getLastCellNum();

                        for (int h = 0; h < cellcount_inner; h++) {

                            // get cell value at the given position [i+1][h]

                            String map_key = sh.getRow(i + 1).getCell(h).getStringCellValue();

                            Cell cell_inner = sh.getRow(i + 2).getCell(h);

                            cell_inner.setCellType(Cell.CELL_TYPE_STRING);

                            String map_value = cell_inner.getStringCellValue();

                            input.put(map_key, map_value);

                        }

                        // incrementing row count to skip next row

                        rowcount = i + 1;

                        break;

                    }

                }

            } catch (NullPointerException e) {

                // System.out.println("caught null pointer exception");

            }

        }

        return input;

    }

}