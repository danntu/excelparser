package polygon;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class CreateXSSFWorkbook {
    public static void main(String[] args) throws IOException,FileNotFoundException {
        Workbook workbook = new XSSFWorkbook();
        FileOutputStream fileOutputStream = new FileOutputStream("src/main/java/polygon/output/workbook.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }
}
