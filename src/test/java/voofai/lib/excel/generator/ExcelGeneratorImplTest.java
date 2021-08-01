package voofai.lib.excel.generator;

import bad.robot.excel.matchers.WorkbookMatcher;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hamcrest.Matcher;
import org.hamcrest.MatcherAssert;
import org.junit.jupiter.api.Test;
import voofai.lib.BaseTest;
import voofai.lib.excel.dto.ExcelData;
import voofai.lib.excel.dto.ExcelSimpleTableData;
import voofai.lib.excel.dto.WorkBookType;

import java.io.*;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;
import java.util.stream.Stream;

class ExcelGeneratorImplTest extends BaseTest {
    private static final String TEMPLATE_PATH = "/excel/template.xlsx";
    private static final String EXPECTED_RESULT_PATH = "/excel/expected_result.xlsx";

    private final ExcelGeneratorImpl excelGenerator = new ExcelGeneratorImpl();

    @Test
    void createWorkbookByTemplateTest() throws IOException {
        try (InputStream template = ExcelGeneratorImplTest.class.getResourceAsStream(TEMPLATE_PATH)) {
            //generate
            ExcelData data = generateData();
            Workbook actual = excelGenerator.createWorkbookByTemplate(WorkBookType.XLSX, data, template);

            //compare expected vs actual
            try (InputStream expected = ExcelGeneratorImplTest.class.getResourceAsStream(EXPECTED_RESULT_PATH)) {
                Matcher<Workbook> matcher = WorkbookMatcher.sameWorkbook(new XSSFWorkbook(expected));
                MatcherAssert.assertThat("Workbooks to be identical", actual, matcher);
            }
        }

    }

    private ExcelData generateData() {
        final ExcelSimpleTableData tableData = new ExcelSimpleTableData("#TABLE&table#");
        final ExcelData excelData = new ExcelData();
        excelData.getSimpleTables().add(tableData);
        excelData.addParam("date", LocalDate.of(2021, 8, 1));

        Stream.of(1, 2, 3, 954)
                .map(this::mapToExcelTableStringData)
                .forEach(tableData::addTableRowData);

        return excelData;
    }

    private Map<String, Object> mapToExcelTableStringData(int postFix) {
        Map<String, Object> map = new HashMap<>();
        map.put("partNumber", String.format("partNumber%d", postFix));
        map.put("name", String.format("name%d", postFix));
        map.put("itemType", String.format("itemType%d", postFix));
        map.put("size", String.format("size%d", postFix));
        return map;
    }

}