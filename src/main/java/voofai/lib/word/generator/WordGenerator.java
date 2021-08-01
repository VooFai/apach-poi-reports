package voofai.lib.word.generator;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import voofai.lib.word.dto.WordData;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Read template file as XWPFDocument and fill cells with required data
 * If there are special symbols in firs cell of row,
 * then create table or range and fill it with required data
 */
public interface WordGenerator {

    /**
     * Read word template file and fill with data from dto
     *
     * @param dto          data to fill template with
     * @param template     word template to make copy and fill with data
     * @param outputStream stream to write result data
     * @throws IOException exception while trying to parse InputStream into XWPFDocument
     */
    void createByTemplate(WordData dto, InputStream template, OutputStream outputStream) throws IOException;

    /**
     * Read word template file and fill with data from dto
     *
     * @param dto      data to fill template with
     * @param template word template to make copy and fill with data
     * @throws IOException exception while trying to parse InputStream into XWPFDocument
     */
    XWPFDocument createDocByTemplate(WordData dto, InputStream template) throws IOException;

    /**
     * Read excell template file and fill with data from dto
     *
     * @param dto          data to fill template with
     * @param templateFile word template to make copy and fill with data
     * @throws IOException exception while trying to parse InputStream into XWPFDocument
     */
    XWPFDocument createDocByTemplate(WordData dto, byte[] templateFile) throws IOException;
}
