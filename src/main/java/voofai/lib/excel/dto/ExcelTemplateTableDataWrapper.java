package voofai.lib.excel.dto;

import lombok.Getter;
import lombok.RequiredArgsConstructor;

import java.util.LinkedList;

/**
 * Данные для эксель-таблицы
 * Навороченные таблицы (key = параметр в excel-ячейке, val = строки рэнджа(объект ExcelTemplateData))
 */
@Getter
@RequiredArgsConstructor
public class ExcelTemplateTableDataWrapper {
    private final String tableName;
    private final LinkedList<ExcelTemplateTableData> tableRowsData = new LinkedList<>();

    public void addTableRowData(ExcelTemplateTableData stringObjectMap) {
        this.tableRowsData.add(stringObjectMap);
    }
}
