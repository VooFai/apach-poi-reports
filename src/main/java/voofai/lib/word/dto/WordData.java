package voofai.lib.word.dto;

import lombok.Getter;
import lombok.RequiredArgsConstructor;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;

/**
 * DTO для генерации Word-документа из существующего шаблона
 */
@Getter
@RequiredArgsConstructor
public class WordData {
    private static final DateTimeFormatter DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("dd.MM.yyyy");

    // key = параметр в excel-ячейке, val = значение
    private final Map<String, String> params = new HashMap<>();

    public WordData addParam(String key, String val) {
        params.put(key, val);
        return this;
    }

    public WordData addParam(String key, Number val) {
        params.put(key, val.toString());
        return this;
    }

    public WordData addParam(String key, LocalDate val) {
        final String valAsString = val != null ? DATE_TIME_FORMATTER.format(val) : "N/A";
        params.put(key, valAsString);
        return this;
    }

    public WordData addAll(Map<String, String> newParams) {
        params.putAll(newParams);
        return this;
    }
}
