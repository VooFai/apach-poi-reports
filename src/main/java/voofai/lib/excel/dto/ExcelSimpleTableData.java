package voofai.lib.excel.dto;

import lombok.Getter;
import lombok.RequiredArgsConstructor;

import java.util.LinkedList;
import java.util.Map;

/**
 * Данные для эксель-таблицы
 * Простые таблицы (key = параметр в excel-ячейке, val = строки таблицы)
 */
@Getter
@RequiredArgsConstructor
public class ExcelSimpleTableData {
    private final String tableName;
    private final LinkedList<Map<String, Object>> tableRowsData = new LinkedList<>();

    public void addTableRowData(Map<String, Object> stringObjectMap) {
        this.tableRowsData.add(stringObjectMap);
    }
}
