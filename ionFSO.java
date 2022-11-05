import java.util.*;
import java.io.*;

public class ionFSO {
    public String openFile(String path) throws FileNotFoundException {
        File file = new File(path);
        Scanner scan = new Scanner(file);
        scan.useDelimiter("\\Z");
        String content = scan.next();
        scan.close();
        return content;
    }

    public Boolean createFile(String path) throws IOException {
        File file = new File(path);

        if (file.createNewFile()) {
            return true;
        } else {
            return false;
        }
    }

    public Boolean writeFile(String path, String value) throws IOException {
        FileWriter file = new FileWriter(path);
        file.write(value);
        file.close();
        return true;
    }

    public Boolean overwriteFile(String path, String value) throws IOException {
        createFile(path);
        writeFile(path, value);
        return true;
    }

    public Boolean deleteFile(String path) {
        File file = new File(path);
        return file.delete();
    }
}