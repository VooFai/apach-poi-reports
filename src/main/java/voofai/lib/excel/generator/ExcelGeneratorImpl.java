package voofai.lib.excel.generator;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import voofai.lib.excel.dto.*;

import javax.validation.constraints.NotNull;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelGeneratorImpl implements ExcelGenerator {

    //template special symbols
    protected static final String TABLE = "#TABLE";
    protected static final String RANGE = "#RANGE";
    protected static final String RANGE_LIST = "#RANGELIST";
    protected static final String BODY = "#Body";
    protected static final String FOOTER = "#Footer";
    protected static final String BODY_FOOTER = "#BodyFooter";
    protected static final String PAGE_BREAKER = "$PageBreaker$";
    protected static final String PAGE_NUM = "$pageNum$";

    //dateTime format
    protected static final String dateTimePattern = "dd.MM.yyyy";
    protected static final DateTimeFormatter localDateFormat = DateTimeFormatter.ofPattern(dateTimePattern);

    @Override
    public Workbook createWorkbookByTemplate(WorkBookType workBookType, ExcelData dto,
                                             InputStream templateFile) throws IOException {
        Workbook workbook = parseTemplateIntoWorkBook(workBookType, templateFile);
        return fillWorkbookWithData(dto, workbook);
    }

    @Override
    public Workbook createWorkbookByTemplate(WorkBookType workBookType, ExcelData dto,
                                             byte[] templateFile) throws IOException {
        try (final ByteArrayInputStream inputStream = new ByteArrayInputStream(templateFile)) {
            Workbook workbook = parseTemplateIntoWorkBook(workBookType, inputStream);
            return fillWorkbookWithData(dto, workbook);
        }
    }

    @Override
    public void createByTemplate(WorkBookType workBookType, ExcelData dto,
                                 InputStream templateFile, OutputStream outputStream) throws IOException {
        Workbook workbook = parseTemplateIntoWorkBook(workBookType, templateFile);
        fillWorkbookWithData(dto, workbook);
        workbook.write(outputStream);
    }


    private Workbook parseTemplateIntoWorkBook(WorkBookType workBookType,
                                               InputStream templateFile) throws IOException {
        return WorkBookType.XLS.equals(workBookType)
                ? new HSSFWorkbook(templateFile)
                : new XSSFWorkbook(templateFile);
    }

    /**
     * !!Side effected code
     * <p>
     * filling entire workbook with data from dto
     *
     * @param dto      data to fill the workbook
     * @param workbook source template
     */
    private Workbook fillWorkbookWithData(@NotNull ExcelData dto, Workbook workbook) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            //choose list
            Sheet sheet = workbook.getSheetAt(i);

            //filling template with data row by row
            int rowNum = sheet.getFirstRowNum();
            while (rowNum <= sheet.getLastRowNum()) {

                Row row = sheet.getRow(rowNum);

                //check if there is a range start
                if (startWithString(row, RANGE)) {
                    String cellVal = row.getCell(0).getStringCellValue();
                    //create range and set rowNum = last range row
                    rowNum = createRange(sheet, rowNum, dto, cellVal);

                    //check if there is table start
                } else if (startWithString(row, TABLE)) {
                    String cellVal = row.getCell(0).getStringCellValue();
                    //create table and set rowNum = last table row
                    rowNum = createTable(sheet, rowNum, dto, cellVal);

                    //simple data = dto.params
                } else {
                    rowNum = fillCells(sheet, rowNum, dto.getParams(), dto);
                }
            }

            //delete unnecessary data (range start\end, table start\end, etc.)
            deleteAllTemplates(sheet);

            //if there is a PAGE_BREAKER - add page breaker for each n rows (and add current row number)
            //splitPage(sheet); //todo just uncomment if you need it (expensive thing)
        }
        return workbook;
    }


    /**
     * Fill cells with data from {@param map}
     * create new rows for tables\ranges if needed
     *
     * @param sheet  current sheet
     * @param rowNum current rowNum
     * @param map    plain data
     * @param dto    data
     * @return next row num (including all subtables rows)
     */

    private int fillCells(Sheet sheet, int rowNum, Map<String, Object> map, ExcelData dto) {
        SimpleDateFormat format = new SimpleDateFormat(dateTimePattern);

        Row row = sheet.getRow(rowNum);
        if (row == null) {
            return rowNum + 1;
        }

        Iterator<Cell> iterator = row.cellIterator();

        // for each cell of current row
        while (iterator.hasNext()) {
            Cell cell = iterator.next();

            //breakpoint
            if (!hasKey(cell)) {
                continue;
            }

            String cellVal = cell.getStringCellValue();

            //check if there are subTables to create
            if (cell.getColumnIndex() == 0 && cellVal.startsWith(RANGE)) {
                //create range (return last row number)
                return createRange(sheet, rowNum, dto, cellVal);

                //create table (return last row number)
            } else if (cell.getColumnIndex() == 0 && cellVal.startsWith(TABLE)) {
                return createTable(sheet, rowNum, dto, cellVal);

                //plain data - just take data from map
            } else {
                String key = cellVal.substring(1, cellVal.length() - 1);
                if (!map.containsKey(key)) {
                    cell.setCellValue("");

                    //set value based on cell type
                } else {
                    Object param = map.get(key);
                    if (param == null) {
                        cell.setCellValue("");
                    } else if (param instanceof Date) {
                        cell.setCellValue(format.format(param));
                    } else if (param instanceof LocalDate) {
                        cell.setCellValue(localDateFormat.format((LocalDate) param));
                    } else if (param instanceof Double) {
                        cell.setCellValue((Double) param);
                    } else if (param instanceof BigDecimal) {
                        BigDecimal bigDecimal = (BigDecimal) param;
                        cell.setCellValue(bigDecimal.doubleValue());
                    } else {
                        cell.setCellValue(param.toString());
                        if (cell.getHyperlink() != null && cell.getHyperlink().getType() == HyperlinkType.URL) {
                            cell.getHyperlink().setAddress(param.toString());
                        }
                    }
                }
            }
        }
        return rowNum + 1;
    }

    /**
     * Check if cell contains specific data
     * (telling that we should fill some data or just skip this cell)
     */
    private boolean hasKey(Cell cell) {
        if (cell == null || !CellType.STRING.equals(cell.getCellType())) {
            return false;
        }
        String cellVal = cell.getStringCellValue();
        return (cellVal.startsWith("#") && cellVal.endsWith("#"));
    }

    /**
     * Check that first cell value equals key value
     *
     * @param row current row
     * @param key key to find
     * @return boolean
     */
    private boolean hasString(Row row, String key) {
        if (row != null) {
            Cell firstCell = row.getCell(0);
            return firstCell != null
                    && CellType.STRING.equals(firstCell.getCellType())
                    && firstCell.getStringCellValue().equals(key);
        }
        return false;
    }

    /**
     * Get data from map with type conversion
     */
    private String getParam(String key, Map<String, Object> map) {
        if (!map.containsKey(key)) {
            return "";
        }
        Object param = map.get(key);
        if (param == null) {
            return "";
        }
        if (param instanceof Date) {
            SimpleDateFormat format = new SimpleDateFormat(dateTimePattern);
            return format.format(param);
        }
        if (param instanceof LocalDate) {
            return localDateFormat.format((LocalDate) param);
        }
        if (param instanceof Double) {
            DecimalFormat df = new DecimalFormat("#.#####");
            DecimalFormatSymbols sym = DecimalFormatSymbols.getInstance();
            sym.setDecimalSeparator(',');
            df.setDecimalFormatSymbols(sym);
            return df.format(param);
        }
        if (param instanceof BigDecimal) {
            BigDecimal bigDecimal = (BigDecimal) param;
            DecimalFormat df = new DecimalFormat("#.#####");
            DecimalFormatSymbols sym = DecimalFormatSymbols.getInstance();
            sym.setDecimalSeparator(',');
            df.setDecimalFormatSymbols(sym);
            return df.format(bigDecimal.doubleValue());
        } else {
            return param.toString();
        }
    }

    /**
     * Копирует строки между строками с значением cellVal. Заполняет скопированные строки.
     *
     * @param sheet   лист
     * @param rowNum  номер строки
     * @param dto     данные
     * @param cellVal значение в ячейке (из него берется название range)
     * @return номер строки окончания range
     */
    private int createTable(Sheet sheet, int rowNum, ExcelData dto, String cellVal) {

        //пропускаем строку "идентификатор начала"
        rowNum++;
        //шаблон строк
        List<Row> sourceRows = new ArrayList<>();
        while (rowNum < sheet.getLastRowNum() && !hasString(sheet.getRow(rowNum), cellVal)) {
            sourceRows.add(sheet.getRow(rowNum));
            rowNum++;
        }
        //пропускаем строку "идентификатор окончания"
        rowNum++;

        Optional<LinkedList<Map<String, Object>>> data = dto.getSimpleTables().stream()
                .filter(range -> range.getTableName().equals(cellVal))
                .findFirst()
                .map(ExcelSimpleTableData::getTableRowsData);

        //условия продолжения
        if (sourceRows.isEmpty() || !data.isPresent()) {
            return rowNum;
        }

        //создание таблиц из шаблонов для каждого dataItem
        for (Map<String, Object> dataItem : data.get()) {
            int startRow = rowNum;
            rowNum = copyRowsAndFillData(dataItem, sheet, startRow, dto, sourceRows);
        }

        return rowNum;
    }

    private int copyRowsAndFillData(Map<String, Object> dataItem, Sheet sheet, int rowNum, ExcelData dto, List<Row> sourceRows) {
        copyRows(sheet, sourceRows, rowNum);

        //Заполнение данными
        for (int index = 0; index < sourceRows.size(); index++) {
            //вложенные таблицы
            if (startWithString(sheet.getRow(rowNum), TABLE)) {
                index = decreaseNumberOfRowsToFill(sheet, rowNum, sourceRows.size(), index);
            }
            // заполнение данными (приравниваем rowNum к последней незаполненной строке)
            rowNum = fillCells(sheet, rowNum, dataItem, dto);
        }
        return rowNum;
    }

    /**
     * Залипуха, связанная с вложенными таблицами
     */
    private int decreaseNumberOfRowsToFill(Sheet sheet, int rowNum, int rowsSize, int index) {
        //название вложенной таблицы
        String stopWordInner = sheet.getRow(rowNum).getCell(0).getStringCellValue();
        //уменьшаем количество строк для fillRange на длину вложенного range
        int counter = 1;
        while (!hasString(sheet.getRow(rowNum + counter), stopWordInner) && counter <= rowsSize) {
            counter++;
        }
        return index + counter + 2;
    }

    /**
     * Копирует строки между двумя ячейками с значением cellVal. Создает range на основе скопированных строк.
     *
     * @param sheet       лист
     * @param startRowNum номер строки
     * @param dto         данные
     * @param cellVal     значение в ячейке (из него берется название range)
     * @return номер строки окончания range
     */
    private int createRange(Sheet sheet, int startRowNum, ExcelData dto, String cellVal) {
        //пропускаем строку "идентификатор начала"
        int rowNum = startRowNum + 1;

        //шаблон строк
        ExcelTemplateTableRows template = createTemplate(sheet, rowNum, cellVal);
        //пролистываем шаблон
        while (rowNum < sheet.getLastRowNum() && !hasString(sheet.getRow(rowNum), cellVal)) {
            rowNum++;
        }
        rowNum++;

        //создание нескольких таблиц из шаблона
        if (cellVal.startsWith(RANGE_LIST)) {

            Set<ExcelTemplateTableDataWrapper> tables = dto.getExtendedTables()
                    .stream()
                    .filter(range -> range.getTableName().startsWith(RANGE_LIST))
                    .collect(Collectors.toSet());

            for (ExcelTemplateTableDataWrapper dataItem : tables) {
                rowNum = fillRange(sheet, template, dto, dataItem.getTableName(), rowNum);
            }

            //создание одной таблицы
        } else {
            rowNum = fillRange(sheet, template, dto, cellVal, rowNum);
        }

        return rowNum;
    }

    /**
     * Заполнение шаблона строк данными
     *
     * @param sheet       лист
     * @param template    шаблон строк
     * @param dto         данные
     * @param mapKey      ключ нужной мапы
     * @param startRowNum первая строка для заполнения
     * @return номер последней незаполненной строки
     */
    private int fillRange(Sheet sheet, ExcelTemplateTableRows template, ExcelData dto, String mapKey, int startRowNum) {
        int rowNum = startRowNum;

        Optional<LinkedList<ExcelTemplateTableData>> tables = dto.getExtendedTables().stream()
                .filter(range -> range.getTableName().equals(mapKey))
                .findFirst()
                .map(ExcelTemplateTableDataWrapper::getTableRowsData);

        List<Row> sourceRows = template.getBody();

        //условия продолжения
        if (sourceRows.isEmpty() || !tables.isPresent()) {
            return rowNum;
        }

        //создание таблиц из щаблонов
        for (ExcelTemplateTableData dataItem : tables.get()) {

            //создание header
            if (!dataItem.getHeader().isEmpty() && !template.getHeader().isEmpty()) {
                copyRows(sheet, template.getHeader(), rowNum); //копии строк шаблона
                for (int i = 0; i < template.getHeader().size(); i++) {
                    rowNum = fillCells(sheet, rowNum, dataItem.getHeader(), dto); // заполнение данными (приравниваем rowNum к последней незаполненной строке)
                }
            }

            //создание body
            for (Map<String, Object> map : dataItem.getBody()) {
                int startRow = rowNum;
                rowNum = copyRowsAndFillData(map, sheet, startRow, dto, sourceRows);
            }

            //создание bodyFooter
            if (!dataItem.getBodyFooter().isEmpty() && !template.getBodyFooter().isEmpty()) {
                for (Map<String, Object> map : dataItem.getBodyFooter()) {
                    copyRows(sheet, template.getBodyFooter(), rowNum); //копии строк шаблона
                    for (int i = 0; i < template.getBodyFooter().size(); i++) {
                        rowNum = fillCells(sheet, rowNum, map, dto); // заполнение данными (приравниваем rowNum к последней незаполненной строке)
                    }
                }
            }

            //создание footer
            if (!dataItem.getFooter().isEmpty() && !template.getFooter().isEmpty()) {
                copyRows(sheet, template.getFooter(), rowNum); //копии строк шаблона
                for (int i = 0; i < template.getFooter().size(); i++) {
                    rowNum = fillCells(sheet, rowNum, dataItem.getFooter(), dto); // заполнение данными (приравниваем rowNum к последней незаполненной строке)
                }
            }
        }

        return rowNum;
    }

    /**
     * Возвращает объект ExcelTemplate - шаблоны строк для построения рэнджа
     *
     * @param sheet       лист
     * @param startRowNum номер строки
     * @param stopWord    название рэнджа
     * @return ExcelTemplate
     */
    private ExcelTemplateTableRows createTemplate(Sheet sheet, int startRowNum, String stopWord) {
        int rowNum = startRowNum;
        ExcelTemplateTableRows template = new ExcelTemplateTableRows();
        /* копии строк, пока не найдет стоп-слово или до конца листа */

        //header для range (обязательный !!может быть ноль строк)
        while (rowNum < sheet.getLastRowNum() && !hasString(sheet.getRow(rowNum), BODY + stopWord)) {
            template.addHeader(sheet.getRow(rowNum));
            rowNum++;
        }
        rowNum++; //проматываем строку BODY + stopWord

        //body (обязательный элемент)
        while (bodyEnds(sheet, rowNum, stopWord)) {

            //bodyFooter (необязательный элемент, может не быть вовсе)
            if (hasString(sheet.getRow(rowNum), BODY_FOOTER + stopWord)) {
                rowNum++;
                // копируем, пока не встретим stopWord или footer
                while (bodyFooterEnds(sheet, rowNum, stopWord)) {
                    template.addBodyFooter(sheet.getRow(rowNum));
                    rowNum++;
                }

                //Footer (необязательный элемент, может не быть вовсе)
            } else if (hasString(sheet.getRow(rowNum), FOOTER + stopWord)) {
                rowNum++;
                // копируем, пока не встретим stopWord
                while (bodyEnds(sheet, rowNum, stopWord)) {
                    template.addFooter(sheet.getRow(rowNum));
                    rowNum++;
                }

                //Body
            } else {
                template.addBody(sheet.getRow(rowNum));
                rowNum++;
            }
        }

        return template;
    }

    /**
     * Окончание body шаблона
     */
    private boolean bodyEnds(Sheet sheet, int rowNum, String stopWord) {
        return rowNum < sheet.getLastRowNum()
                && !hasString(sheet.getRow(rowNum), stopWord);
    }

    /**
     * Окончание bodyFooter шаблона
     */
    private boolean bodyFooterEnds(Sheet sheet, int rowNum, String stopWord) {
        return rowNum < sheet.getLastRowNum()
                && !hasString(sheet.getRow(rowNum), stopWord)
                && !hasString(sheet.getRow(rowNum), FOOTER + stopWord);
    }

    /**
     * удаление строки
     *
     * @param sheet  лист
     * @param rowNum номер строки
     */
    private void dellRow(Sheet sheet, int rowNum) {
        sheet.removeRow(sheet.getRow(rowNum));

        //залипуха с номером последней строки в экселе (причина = считаем с 0)
        int maxWorkSheetRowNum = sheet.getWorkbook().getSpreadsheetVersion().getMaxRows() - 1; //начинается с нуля (max для XLSX = 1048575)
        int lastRow = maxWorkSheetRowNum == sheet.getLastRowNum() ? maxWorkSheetRowNum - 1 : sheet.getLastRowNum();

        sheet.shiftRows(rowNum + 1, lastRow, -1);
    }

    /**
     * Удаление шаблонов для range
     *
     * @param sheet лист
     */
    private void deleteAllTemplates(Sheet sheet) {
        for (int rowNum = sheet.getFirstRowNum(); rowNum < sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);

            //проверяем есть ли идентификатор начала range
            if (startWithString(row, RANGE) || startWithString(row, TABLE)) {
                int endRowNum = rowNum;
                //название range
                String cellVal = row.getCell(0).getStringCellValue();

                //проматываем до окончания range
                endRowNum++;
                while (endRowNum < sheet.getLastRowNum() && !hasString(sheet.getRow(endRowNum), cellVal)) {
                    endRowNum++;
                }
                if (endRowNum < sheet.getLastRowNum()) {
                    endRowNum++;
                }

                //удаление range
                dellRange(sheet, rowNum, endRowNum);
            }
        }
    }

    private void copyRow(Sheet sheet, Row sourceRow, int destinationRowNum) {
        copyRow(sheet, sourceRow, destinationRowNum, true);
    }

    /**
     * Копирование строки
     *
     * @param sheet             лист назначения
     * @param sourceRow         source
     * @param destinationRowNum номер строки, в которую копировать
     */
    private void copyRow(Sheet sheet, Row sourceRow, int destinationRowNum, boolean copyMergedRanges) {
        Row newRow = sheet.getRow(destinationRowNum);

        // Если есть - шифт строк. Если нет - создание строки
        if (newRow != null) {
            //залипуха с номером последней строки в экселе (причина = считаем с 0)
            int maxWorkSheetRowNum = sheet.getWorkbook().getSpreadsheetVersion().getMaxRows() - 1; //начинается с нуля (max для XLSX = 1048575)
            int lastRow = maxWorkSheetRowNum == sheet.getLastRowNum()
                    ? maxWorkSheetRowNum - 1
                    : sheet.getLastRowNum();

            sheet.shiftRows(destinationRowNum, lastRow, 1);
            newRow = sheet.createRow(destinationRowNum);
        } else {
            newRow = sheet.createRow(destinationRowNum);
        }

        Iterator<Cell> cellIterator = sourceRow.cellIterator();
        while (cellIterator.hasNext()) {
            Cell oldCell = cellIterator.next();
            Cell newCell = newRow.createCell(oldCell.getColumnIndex());

            // style
            newCell.setCellStyle(oldCell.getCellStyle());
            newCell.setCellComment(oldCell.getCellComment());
            newCell.setCellType(oldCell.getCellType());

            //hyperlink
            Hyperlink hyperlink = oldCell.getHyperlink();
            if (hyperlink != null) {
                final CreationHelper creationHelper = sheet.getWorkbook().getCreationHelper();
                Hyperlink link = creationHelper.createHyperlink(hyperlink.getType());
                link.setAddress(hyperlink.getAddress());
                link.setLabel(hyperlink.getLabel());
                newCell.setHyperlink(link);
            }

            // value
            switch (oldCell.getCellType()) {
                case BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
                case BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                case NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                default:
                    newCell.setBlank();
            }
        }

        // Копирование merged Regions
        if (copyMergedRanges) {
            sourceRow.getSheet().getMergedRegions().stream()
                    .filter(merged -> merged.getFirstRow() == sourceRow.getRowNum())
                    .forEach(merged -> {
                        int lastRow = destinationRowNum + (merged.getLastRow() - merged.getFirstRow());
                        CellRangeAddress newCellRangeAddress = new CellRangeAddress(
                                destinationRowNum, lastRow, merged.getFirstColumn(), merged.getLastColumn());
                        sheet.addMergedRegion(newCellRangeAddress);
                    });
        }
    }

    /**
     * Проверка налиия ключа-идентификатора данных в первой ячейке (startsWith)
     *
     * @param row строка
     * @param key искомый ключ
     * @return boolean
     */
    private boolean startWithString(Row row, String key) {
        if (row != null) {
            Cell firstCell = row.getCell(0);
            return firstCell != null
                    && CellType.STRING.equals(firstCell.getCellType())
                    && firstCell.getStringCellValue().startsWith(key);
        }
        return false;
    }

    /**
     * Если есть строка с пометкой PAGE_BREAKER - разбить на страницы
     * пример строки-идентификатора
     * $PageBreaker$33-2
     * , где 33 - количество строк, через которые вставлять разрыв страницы
     * 2 - номер первой страницы
     *
     * @param sheet лист
     */
    private void splitPage(Sheet sheet) {
        int rowNum = sheet.getFirstRowNum();
        while (rowNum < sheet.getLastRowNum()) {
            Row row = sheet.getRow(rowNum);

            //признак необходимости постраничной разбивки
            if (startWithString(row, PAGE_BREAKER)) {
                String cellVal = row.getCell(0).getStringCellValue();
                String restData = cellVal.substring(PAGE_BREAKER.length());
                String[] params = restData.split("-");

                //получаем длину строк в листе и номер первого листа
                int pageLength = Integer.parseInt(params[0]);
                int firstPage = Integer.parseInt(params[1]);

                //постраничная разбивка
                splitPage(sheet, rowNum, pageLength, firstPage);

                break;
            }
            rowNum++;
        }
    }

    /**
     * Разбивка на страницы начиная с указанной строки
     *
     * @param sheet            лист
     * @param templateStartRow начальна трока
     */
    private void splitPage(Sheet sheet, int templateStartRow, int pageLength, int startPage) {
        int pageNum = startPage;
        int rowNum = templateStartRow;

        //создаем шаблон с строками, копируемыми в конец каждого листа (номер листа к примеру)
        List<Row> template = new ArrayList<>();
        rowNum++;
        while (rowNum < sheet.getLastRowNum() && !startWithString(sheet.getRow(rowNum), PAGE_BREAKER)) {
            template.add(sheet.getRow(rowNum));
            rowNum++;
        }
        rowNum++;

        int templateEndRow = rowNum;

        //пролистываем на длину pageLength
        rowNum = rowNum + 1 + pageLength - template.size();

        while ((sheet.getLastRowNum() - rowNum) > pageLength) { //пока до конца листа не останется строк меньше, чем pageLength
            //копирование строк из шаблона
            copyRows(sheet, template, rowNum);

            //проставляем номер страницы
            for (int i = 0; i <= template.size(); i++) {
                if (sheet.getRow(rowNum) == null) {
                    break;
                }
                Iterator<Cell> iterator = sheet.getRow(rowNum).cellIterator();
                while (iterator.hasNext()) {
                    Cell cell = iterator.next();
                    if (cell != null
                            && CellType.STRING.equals(cell.getCellType())
                            && cell.getStringCellValue().equals(PAGE_NUM)
                    ) {
                        //задание номера текущего листа
                        pageNum++;
                        cell.setCellValue(pageNum);
                    }
                }
                rowNum++;
            }

            //добавляем разрыв страниц
            sheet.setRowBreak(rowNum - 2);

            //пролистываем на длину pageLength
            rowNum = rowNum + pageLength;
        }

        //удаляем шаблон
        dellRange(sheet, templateStartRow, templateEndRow);
    }

    /**
     * Копирование строк исходных строк в данный sheet, начиная с строки startRow
     *
     * @param sheet      лист назначения
     * @param sourceRows source
     * @param startRow   первая строка назначения
     */
    private void copyRows(Sheet sheet, List<Row> sourceRows, int startRow) {
        //копия строк
        int destinationRowNum = startRow;
        for (Row row : sourceRows) {
            if (row == null) {
                continue;
            }
            copyRow(sheet, row, destinationRowNum, false);
            destinationRowNum++;
        }

        //создание копий merged regions
        destinationRowNum = startRow;
        for (Row row : sourceRows) {
            if (row == null) {
                continue;
            }
            for (CellRangeAddress merged : row.getSheet().getMergedRegions()) {
                if (merged.getFirstRow() == row.getRowNum()) {
                    int lastRow = destinationRowNum + (merged.getLastRow() - merged.getFirstRow());
                    CellRangeAddress newCellRangeAddress = new CellRangeAddress(
                            destinationRowNum, lastRow, merged.getFirstColumn(), merged.getLastColumn());
                    sheet.addMergedRegion(newCellRangeAddress);
                }
            }
            destinationRowNum++;
        }
    }

    /**
     * удаление строк
     *
     * @param sheet    лист
     * @param startRow первая строка рэнжда
     * @param endRow   последняя
     */
    private void dellRange(Sheet sheet, int startRow, int endRow) {
        //удаление merged regions
        for (int rowNum = startRow; rowNum < endRow; rowNum++) {
            int rowNumToDell = 0;
            for (CellRangeAddress merged : sheet.getMergedRegions()) {
                if (merged.getFirstRow() == rowNum) {
                    sheet.removeMergedRegion(rowNumToDell);
                } else {
                    rowNumToDell++;
                }
            }
        }
        //очистка содержимого
        for (int i = startRow; i < endRow; i++) {
            sheet.removeRow(sheet.getRow(i));
        }
        //шифт вверх
        sheet.shiftRows(endRow, sheet.getLastRowNum(), startRow - endRow);
    }
}
