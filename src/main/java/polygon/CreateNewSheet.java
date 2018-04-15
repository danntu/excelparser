package polygon;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class CreateNewSheet {
    public static void main(String[] args) throws IOException,FileNotFoundException{
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("Первая вкладка1");
        Sheet sheet2 = workbook.createSheet("Вторая вкладка2");
        Sheet sheet3 = workbook.createSheet(WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]"));
        FileOutputStream fileOutputStream = new FileOutputStream("src/main/java/polygon/output/workbook.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }
}
