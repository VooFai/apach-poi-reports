package voofai.lib.word.generator;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;
import voofai.lib.BaseTest;
import voofai.lib.word.dto.WordData;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;

class WordGeneratorImplTest extends BaseTest {

    private static final String TEMPLATE_PATH = "/word/template.docx";
    private static final String EXPECTED_RESULT_PATH = "/word/expected_result.docx";

    private final WordGeneratorImpl wordGenerator = new WordGeneratorImpl();

    @Test
    @Disabled("Найти простой способ быстро сравнить содержимое")
    void createDocByTemplateTest() throws IOException {
        File tempFile = newTempFile();

        try (InputStream template = WordGeneratorImplTest.class.getResourceAsStream(TEMPLATE_PATH)) {
            //generate
            WordData data = generateData();
            XWPFDocument actual = wordGenerator.createDocByTemplate(data, template);
            actual.write(new FileOutputStream(tempFile));

            //compare expected vs actual
            //URL resource = WordGeneratorImplTest.class.getResource(EXPECTED_RESULT_PATH);
            //File expected = new File(resource.toURI().getPath());

            //assertTrue(FileUtils.contentEquals(tempFile, expected));
        }
    }

    private WordData generateData() {
        return new WordData()
                .addParam("toolSerial", 9771232)
                .addParam("dateRelease", LocalDate.of(2021, 8, 1))
                .addParam("adr_dateCal", LocalDate.of(3021, 8, 2))
                .addParam("adr_serial", "serial#12345")
                .addParam("toolSize", "_val_999_");
    }
}