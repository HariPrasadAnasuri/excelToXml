package com.bvirtoso.excelToXml.service

import org.apache.commons.lang3.StringUtils
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service

@Service
class ExcelFileService {
    Object readExcel(String FILE_NAME, int sheetIndex) {
        Map<String, String> map = new HashMap<>()
        def xmlInput = [];
        try {
            File file = new File(FILE_NAME)
            FileInputStream excelFile = new FileInputStream(file)
            XSSFWorkbook workbook = new XSSFWorkbook(excelFile)
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex)

            XSSFRow firstRow = sheet.getRow(0)
            int minColumnIndex = firstRow.getFirstCellNum()
            int maxnColumnIndex = sheet.getRow(0).getLastCellNum()
            Map<String, Integer> columnNameMap = new HashMap<>()
            (minColumnIndex..maxnColumnIndex-1).each {
                index ->
                    XSSFCell cell = firstRow.getCell(index)
                    columnNameMap.put(cell.getStringCellValue(), cell.getColumnIndex())
            }

            Iterator<Row> iterator = sheet.iterator();
            int rowCount = 0;
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                if(!checkIfRowIsEmpty(currentRow)) {
                    def input = [
                            'name'   : '',
                            'age'    : '',
                            'address': [
                                    'housenumber': '',
                                    'street'     : '',
                                    'district'   : '',
                                    'town'       : '',
                                    'zipcode'    : ''
                            ]

                    ]

                    if (rowCount > 0) {
                        columnNameMap.each {
                            columnName, columnIndex ->
                                Cell currentCell = currentRow.getCell(columnIndex)
                                //currentCell.setCellType(CellType.STRING)
                                //currentCell.setCellFormula(String.cl)
                                String[] properties = columnName.split("_")
                                if (properties[0].length() > 0) {
                                    if (properties.length < 2) {
                                        input[properties[0]] = currentCell.toString().trim()
                                    } else {
                                        input[properties[0]][properties[1]] = currentCell.toString().trim()
                                    }
                                }

                        }
                        xmlInput.add(input)
                    }
                    rowCount++
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace()
        } catch (IOException e) {
            e.printStackTrace()
        }
        return xmlInput
    }

    private boolean checkIfRowIsEmpty(Row row) {
        if (row == null) {
            return true;
        }
        if (row.getLastCellNum() <= 0) {
            return true;
        }
        for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
            Cell cell = row.getCell(cellNum);
            if (cell != null && cell.getCellTypeEnum() != CellType.BLANK && StringUtils.isNotBlank(cell.toString())) {
                return false;
            }
        }
        return true;
    }
}
