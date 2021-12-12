package voofai.lib.word.generator;

import org.apache.poi.xwpf.usermodel.*;
import voofai.lib.word.dto.WordData;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Predicate;

/**
 * Read template file as XWPFDocument and fill cells with required data
 * If there are special symbols in firs cell of row,
 * then create table or range and fill it with required data
 */
public class WordGeneratorImpl implements WordGenerator {

    //template special symbols
    protected static final String STARTER = "#*";
    protected static final String END = "*#";

    //predicates
    private final Predicate<XWPFParagraph> filterOnlyParagraphWithNotEmptyRuns = p -> !p.getRuns().isEmpty();
    private final Predicate<XWPFParagraph> filterOnlyParagraphWithStarterOnFirstRunString = p -> {
        final XWPFRun firstRun = p.getRuns().get(0);
        if (firstRun != null) {
            final String text = firstRun.getText(0);
            return text != null && text.startsWith(STARTER);
        }
        return false;
    };
    private final Predicate<XWPFTableCell> filterOnlyCellsWithStartEnd = cell -> {
        String cellText = cell.getText();
        return cellText.startsWith(STARTER) && cellText.endsWith(END);
    };

    //consumers
    private final Consumer<XWPFParagraph> removeStarterFromFirstRunString = p -> {
        final XWPFRun firstRun = p.getRuns().get(0);
        final String text = firstRun.getText(0);
        final String textNoStarter = text.substring(STARTER.length());
        firstRun.setText(textNoStarter, 0);
    };

    //biConsumers
    private final BiConsumer<XWPFParagraph, WordData> processParagraph = (paragraph, dto) -> {
        int searchPos = 0;
        TextSegment startRange = paragraph.searchText(STARTER, new PositionInParagraph(searchPos, 0, 0));
        TextSegment endRange = paragraph.searchText(END, new PositionInParagraph(searchPos, 0, 0));
        while (startRange != null && endRange != null) {
            int startPos = startRange.getBeginRun();
            int runsToDellNum = endRange.getEndRun() - startRange.getBeginRun();

            //извлечение текста между STARTER & END
            TextSegment textSegment = new TextSegment(startRange.getBeginPos(), endRange.getEndPos());
            String segmentText = paragraph.getText(textSegment);

            //удалить все run, которые относятся к STARTER
            for (int i = 0; i < runsToDellNum; i++) {
                paragraph.removeRun(startPos);
            }

            //run with oldText
            XWPFRun oldRun = paragraph.getRuns().get(startPos);
            final String paramVal = removeStartEnd(segmentText);
            final String newVal = dto.getParams().getOrDefault(paramVal, "");
            oldRun.setText(newVal, 0);

            searchPos = startPos + 1;
            startRange = paragraph.searchText(STARTER, new PositionInParagraph(searchPos, 0, 0));
            endRange = paragraph.searchText(END, new PositionInParagraph(searchPos, 0, 0));
        }
    };

    /**
     * Обрезает STARTER и END
     */
    private String removeStartEnd(String str) {
        return str.substring(STARTER.length(), str.length() - END.length());
    }

    /**
     * Читает ворд-файл в качестве шаблона и заполняет данными из dto
     * <p>
     * При этом, в шаблоне ищется параграф, который начинается с символа {@value STARTER}
     * Далее, в этом параграфе ищутся все сегменты, начинающиеся на {@value STARTER} и оканчивающиеся на {@value END}
     * Берутся текстовые значения этих сегментов (без символов {@value STARTER} и {@value END})
     * Эти текстовые значения будут ключами к мапе {@link WordData},
     * из которой берутся значения и подставляются вместо текста в данном сегменте.
     * <p>
     * Для таблиц проверяются все ячейки во всех строках на соответствие:
     * строка начинается на {@value STARTER} и оканчивающиеся на {@value END}.
     * Далее та же логика с сегментами.
     *
     * @param doc исходный шаблон
     * @param dto данные
     * @return Workbook
     */
    private XWPFDocument generateWord(XWPFDocument doc, WordData dto) {

        //plain text
        doc.getParagraphs().stream()
                .filter(filterOnlyParagraphWithNotEmptyRuns)
                .filter(filterOnlyParagraphWithStarterOnFirstRunString)
                .peek(removeStarterFromFirstRunString)
                .forEach(paragraph -> processParagraph.accept(paragraph, dto));

        //tables
        doc.getTables().forEach(table -> table.getRows().forEach(row ->
                row.getTableCells().stream()
                        .filter(filterOnlyCellsWithStartEnd)
                        .forEach(cell -> cell.getParagraphs().forEach(p -> processParagraph.accept(p, dto)))));

        return doc;
    }

    @Override
    public XWPFDocument createDocByTemplate(WordData dto, InputStream template) throws IOException {
        XWPFDocument doc = new XWPFDocument(template);
        return generateWord(doc, dto);
    }

    @Override
    public XWPFDocument createDocByTemplate(WordData dto, byte[] templateFile) throws IOException {
        try (final ByteArrayInputStream inputStream = new ByteArrayInputStream(templateFile)) {
            XWPFDocument doc = new XWPFDocument(inputStream);
            return generateWord(doc, dto);
        }
    }

    @Override
    public void createByTemplate(WordData dto, InputStream template, OutputStream outputStream) throws IOException {
        XWPFDocument doc = new XWPFDocument(template);
        generateWord(doc, dto);
        doc.write(outputStream);
    }
}
