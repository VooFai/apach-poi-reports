package voofai.lib.excel.dto;

import lombok.Builder;
import lombok.Getter;
import lombok.RequiredArgsConstructor;

import java.util.HashMap;
import java.util.LinkedList;
import java.util.Map;

/**
 * Данные, которыми будут заполнятся строки из {@link ExcelTemplateTableRows}
 * Суть процесса: построчно считываются строки из размеченного эксель-документа,
 * копируются в {@link ExcelTemplateTableRows},
 * копируются в нужные места в эксель-документ,
 * заполняются значениями из этого класса {@link ExcelTemplateTableData}
 */
@Getter
@RequiredArgsConstructor
public class ExcelTemplateTableData {
    private final Map<String, Object> header = new HashMap<>();
    private final LinkedList<Map<String, Object>> body = new LinkedList<>();
    private final LinkedList<Map<String, Object>> bodyFooter = new LinkedList<>();
    private final Map<String, Object> footer = new HashMap<>();

    public void addHeader(String key, Object data) {
        this.header.put(key, data);
    }

    public void addBody(Map<String, Object> data) {
        this.body.add(data);
    }

    public void addFooter(String key, Object data) {
        this.footer.put(key, data);
    }

    public void addBodyFooter(Map<String, Object> data) {
        this.bodyFooter.add(data);
    }
}
