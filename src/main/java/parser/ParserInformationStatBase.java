package parser;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import javax.xml.transform.TransformerException;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ParserInformationStatBase {
    public static void main(String[] args) throws IOException, FileNotFoundException, TransformerException {
        FileInputStream fileInputStream = new FileInputStream("src/main/java/polygon/input/ВВП по кварталам рус.xls");
        File file = new File("/home/mdaniyar/mygit/excelparser/src/main/java/polygon/input/ВВП по кварталам рус.xls");
        BufferedWriter writer = new BufferedWriter(new FileWriter("/home/mdaniyar/mygit/excelparser/src/main/java/polygon/output/information.txt"));
        List<String> tabNames = new ArrayList<String>();
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        workbook.getNumberOfSheets();
        writer.write("Название файла = "+file.getName()+"\n");
        writer.write("Количество вкладок = "+workbook.getNumberOfSheets()+"\n");
        for (int i =0; i<workbook.getNumberOfSheets();i++){
            writer.write("Название вкладок = "+workbook.getSheetName(i)+"\n");
            tabNames.add(workbook.getSheetName(i));
        }
        writer.close();
        ConstructorsForXMLStatbase xml =  new ConstructorsForXMLStatbase();
        xml.ParamLangXML();
        xml.WriteParamXML(tabNames);
    }
}
