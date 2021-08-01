package voofai.lib;

import org.apache.commons.io.IOUtils;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.assertTrue;

public abstract class BaseTest {

    @TempDir
    protected File tempDir;

    protected File newTempFile() {
        return new File(tempDir, UUID.randomUUID().toString() + ".tmp");
    }

    protected static void assertStreamsEquals(InputStream expected, InputStream actual) throws IOException {
        assertTrue(IOUtils.contentEquals(expected, actual));
    }

}
