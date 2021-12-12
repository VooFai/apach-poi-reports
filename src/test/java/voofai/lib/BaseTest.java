package voofai.lib;

import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.util.UUID;

public abstract class BaseTest {

    @TempDir
    protected File tempDir;

    protected File newTempFile() {
        return new File(tempDir, UUID.randomUUID().toString() + ".tmp");
    }

}
