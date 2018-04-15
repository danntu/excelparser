package parser;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.xml.transform.TransformerException;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ParserInformationReady {
    public static void main(String[] args) throws IOException, FileNotFoundException, TransformerException {
        FileInputStream fileInputStream = new FileInputStream("src/main/java/polygon/input/2017_Мухит_KZ_RU_BY_KG_AM.xlsx");
        File file = new File("/home/mdaniyar/mygit/excelparser/src/main/java/polygon/input/2017_Мухит_KZ_RU_BY_KG_AM.xlsx");
        BufferedWriter writer = new BufferedWriter(new FileWriter("/home/mdaniyar/mygit/excelparser/src/main/java/polygon/output/informationReady.txt"));
        List<String> tabNames = new ArrayList<String>();
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        workbook.getNumberOfSheets();
        writer.write("Название файла = "+file.getName()+"\n");
        writer.write("Количество вкладок = "+workbook.getNumberOfSheets()+"\n");
        for (int i =0; i<workbook.getNumberOfSheets();i++){
            writer.write("Название вкладок = "+workbook.getSheetName(i)+"\n");
            tabNames.add(workbook.getSheetName(i));
        }
        writer.close();

        ConstructorsForXMLReady xml =  new ConstructorsForXMLReady();
        xml.ParamLangXML();
        xml.WriteParamXML(tabNames);

    }
}
