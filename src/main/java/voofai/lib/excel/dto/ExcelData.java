package voofai.lib.excel.dto;

import lombok.Getter;
import lombok.RequiredArgsConstructor;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

/**
 * DTO для генерации Excel-документа из существующего шаблона
 */
@Getter
@RequiredArgsConstructor
public class ExcelData {
    private static final DateTimeFormatter DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("dd.MM.yyyy");

    // key = параметр в excel-ячейке, val = значение
    private final Map<String, Object> params = new HashMap<>();

    private final Set<ExcelSimpleTableData> simpleTables = new HashSet<>();

    private final Set<ExcelTemplateTableDataWrapper> extendedTables = new HashSet<>();

    public ExcelData addParam(String key, String val) {
        params.put(key, val);
        return this;
    }

    public ExcelData addParam(String key, Integer val) {
        params.put(key, val);
        return this;
    }

    public ExcelData addParam(String key, LocalDate val) {
        final String valAsString = val != null ? DATE_TIME_FORMATTER.format(val) : "N/A";
        params.put(key, valAsString);
        return this;
    }
}
