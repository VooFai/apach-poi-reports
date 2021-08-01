package voofai.lib.excel.generator;

import org.apache.poi.ss.usermodel.Workbook;
import voofai.lib.excel.dto.ExcelData;
import voofai.lib.excel.dto.WorkBookType;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Read template file as Workbook and fill cells with required data
 * If there are special symbols in firs cell of row,
 * then create table or range and fill it with required data
 */
public interface ExcelGenerator {

    /**
     * Read excell template file and fill with data from dto
     *
     * @param workBookType .xls vs .xlsx file format
     * @param dto          data to fill template with
     * @param templateFile excel template to make copy and fill with data
     * @param outputStream stream to write result data
     * @throws IOException exception while trying to parse InputStream into Workbook
     */
    void createByTemplate(WorkBookType workBookType,
                          ExcelData dto,
                          InputStream templateFile,
                          OutputStream outputStream) throws IOException;

    /**
     * Read excell template file and fill with data from dto
     *
     * @param workBookType .xls vs .xlsx file format
     * @param dto          data to fill template with
     * @param templateFile excel template to make copy and fill with data
     * @return Workbook
     * @throws IOException exception while trying to parse InputStream into Workbook
     */
    Workbook createWorkbookByTemplate(WorkBookType workBookType, ExcelData dto, InputStream templateFile)
            throws IOException;

    /**
     * Read excell template file and fill with data from dto
     *
     * @param workBookType .xls vs .xlsx file format
     * @param dto          data to fill template with
     * @param templateFile excel template to make copy and fill with data
     * @return Workbook
     * @throws IOException exception while trying to parse InputStream into Workbook
     */
    Workbook createWorkbookByTemplate(WorkBookType workBookType, ExcelData dto, byte[] templateFile)
            throws IOException;
}
