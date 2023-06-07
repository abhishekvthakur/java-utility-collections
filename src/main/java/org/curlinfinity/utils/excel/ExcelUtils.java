package org.curlinfinity.utils.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.HashMap;
import java.util.Map;

public class ExcelUtils {

    public Map<String, Integer> getColumnNameIndex(Sheet sheet) {
        return getColumnNameIndex(sheet, 0);
    }

    public Map<String, Integer> getColumnNameIndex(Sheet sheet, int headerRowIndex) {
        Map<String, Integer> result = new HashMap<>();

        Row headerRow = sheet.getRow(headerRowIndex);
        int count = 1;

        for (Cell cell: headerRow) {
            String cellValue = cell.getRichStringCellValue().getString();
            if (cellValue != null && cellValue.trim().length() > 0) {
                result.put(cellValue, count);
            }

            count++;
        }

        return result;
    }

}
