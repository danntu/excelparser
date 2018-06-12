package polygon;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

public class CreateNewSheet {
    public static void main(String[] args) throws IOException,FileNotFoundException{
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("Первая вкладка1");
        Sheet sheet2 = workbook.createSheet("Вторая вкладка2");
        Sheet sheet3 = workbook.createSheet(WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]"));
        Row row = sheet1.createRow(0);
        row.createCell(0).setCellValue(1.1);
        row.createCell(1).setCellValue(new Date());
        row.createCell(2).setCellValue(Calendar.getInstance());
        row.createCell(3).setCellValue("a string");
        row.createCell(4).setCellValue(true);
        row.createCell(5).setCellType(CellType.ERROR);


        FileOutputStream fileOutputStream = new FileOutputStream("src/main/java/polygon/output/workbook.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }
}
