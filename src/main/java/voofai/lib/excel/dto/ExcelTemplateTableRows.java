package voofai.lib.excel.dto;

import lombok.Builder;
import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Row;

import java.util.LinkedList;

/**
 * В этот класс копируются строки из размеченного заранее эксель-документа
 * Далее, обычным копированием создается нужное количество строк в документе
 * и каждая заполняется данными в порядке добавления в лист
 */
@Getter
@RequiredArgsConstructor
public class ExcelTemplateTableRows {
    private final LinkedList<Row> header = new LinkedList<>();
    private final LinkedList<Row> body = new LinkedList<>();
    private final LinkedList<Row> bodyFooter = new LinkedList<>();
    private final LinkedList<Row> footer = new LinkedList<>();

    public void addHeader(Row row) {
        this.header.add(row);
    }

    public void addBody(Row row) {
        this.body.add(row);
    }

    public void addBodyFooter(Row row) {
        this.bodyFooter.add(row);
    }

    public void addFooter(Row row) {
        this.footer.add(row);
    }
}
