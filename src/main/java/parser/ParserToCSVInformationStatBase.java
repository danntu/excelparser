package parser;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ParserToCSVInformationStatBase {

    static private Pattern rxquote = Pattern.compile("\"");
    static private String encodeValue(String value) {
        boolean needQuotes = false;
        if ( value.indexOf(',') != -1 || value.indexOf('"') != -1 ||
                value.indexOf('\n') != -1 || value.indexOf('\r') != -1 )
            needQuotes = true;
        Matcher m = rxquote.matcher(value);
        if ( m.find() ) needQuotes = true; value = m.replaceAll("\"\"");
        if ( needQuotes ) return "\"" + value + "\"";
        else return value;
    }

    public static void main(String[] args) throws IOException {
        FileInputStream fileInputStream = new FileInputStream("src/main/java/polygon/input/ВВП по кварталам рус.xls");
        File file = new File("/home/mdaniyar/mygit/excelparser/src/main/java/polygon/input/ВВП по кварталам рус.xls");
        //BufferedWriter writer = new BufferedWriter(new FileWriter("/home/mdaniyar/mygit/excelparser/src/main/java/polygon/output/information.csv"));
        Workbook wb = new HSSFWorkbook(fileInputStream);

        int index=0;
        int sheetNo = wb.getNumberOfSheets()-1;

        FormulaEvaluator fe = null;
        if ( index < args.length ) {
            fe = wb.getCreationHelper().createFormulaEvaluator();
        }

        DataFormatter formatter = new DataFormatter();
        PrintStream out = new PrintStream(new FileOutputStream("/home/mdaniyar/mygit/excelparser/src/main/java/polygon/output/information.csv"),
                true, "UTF-8");

        byte[] bom = {(byte)0xEF, (byte)0xBB, (byte)0xBF};
        out.write(bom);
        {
            Sheet sheet = wb.getSheetAt(sheetNo);
            for (int r = 0, rn = sheet.getLastRowNum() ; r <= rn ; r++) {
                Row row = sheet.getRow(r);
                if ( row == null ) { out.println(','); continue; }
                boolean firstCell = true;
                for (int c = 0, cn = row.getLastCellNum() ; c < cn ; c++) {
                    Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if ( ! firstCell ) out.print('&');
                    if ( cell != null ) {
                        if ( fe != null ) cell = fe.evaluateInCell(cell);
                        String value = formatter.formatCellValue(cell);
                        if ( cell.getCellTypeEnum() == CellType.FORMULA ) {
                            value = "=" + value;
                        }
                        //out.print(encodeValue(value));

                    }
                    firstCell = false;
                }
                out.println();
            }
        }



    }
}
