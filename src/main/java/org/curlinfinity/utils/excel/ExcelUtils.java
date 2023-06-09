package org.curlinfinity.utils.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
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

    public Map<String, List<Cell>> filterColumns(Sheet sheet, String[] columnNames) {
        Map<String, List<Cell>> result = new HashMap<>();
        for (String columnName: columnNames) {
            result.put(columnName, new ArrayList<>());
        }

        Map<String, Integer> columnNameIndex = getColumnNameIndex(sheet);
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);

            for (Map.Entry<String, Integer> entry: columnNameIndex.entrySet()) {
                String columnName = entry.getKey();
                Integer columnIndex = entry.getValue();

                result.get(columnName).add(row.getCell(columnIndex));
            }
        }

        return result;
    }

}
