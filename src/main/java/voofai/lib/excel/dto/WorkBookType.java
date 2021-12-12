package voofai.lib.excel.dto;

/**
 * Specifies type of Excel data format
 */
public enum WorkBookType {

    /**
     * HSSFWorkbook - ".xls" file format (before 2007 year)
     */
    XLS,

    /**
     * XSSFWorkbook - ".xlsx" file format (starting at 2007 year)
     */
    XLSX
}
