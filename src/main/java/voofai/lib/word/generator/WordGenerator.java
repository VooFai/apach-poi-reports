package voofai.lib.word.generator;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import voofai.lib.word.dto.WordData;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Generates word document out of specific word template
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
     * @return result word POIXMLDocument
     * @throws IOException exception while trying to parse InputStream into XWPFDocument
     */
    XWPFDocument createDocByTemplate(WordData dto, InputStream template) throws IOException;

    /**
     * Read excell template file and fill with data from dto
     *
     * @param dto          data to fill template with
     * @param templateFile word template to make copy and fill with data
     * @return result word POIXMLDocument
     * @throws IOException exception while trying to parse InputStream into XWPFDocument
     */
    XWPFDocument createDocByTemplate(WordData dto, byte[] templateFile) throws IOException;
}
