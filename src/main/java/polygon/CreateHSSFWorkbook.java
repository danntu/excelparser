package polygon;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class CreateHSSFWorkbook {
    public static void main(String[] args) throws FileNotFoundException,IOException{
        Workbook workbook = new HSSFWorkbook();
        FileOutputStream fileOutputStream = new FileOutputStream("src/main/java/polygon/output/workbook.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }
}
