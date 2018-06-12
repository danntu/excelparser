package parser;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import javax.xml.transform.TransformerException;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

import static java.lang.System.out;

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
        List<String> rowA = VvpByQuarter.readDataA(workbook,tabNames);
        List<String> rowB = VvpByQuarter.readDataB(workbook,tabNames);
        List<String> rowC = VvpByQuarter.readDataC(workbook,tabNames);
        List<String> rowD = VvpByQuarter.readDataD(workbook,tabNames);
        List<String> rowE = VvpByQuarter.readDataE(workbook,tabNames);
        List<String> rowF = VvpByQuarter.readDataF(workbook,tabNames);
        List<String> rowH = VvpByQuarter.readDataH(workbook,tabNames);
        List<String> rowI = VvpByQuarter.readDataI(workbook,tabNames);
        List<String> rowJ = VvpByQuarter.readDataJ(workbook,tabNames);
        List<String> rowK = VvpByQuarter.readDataK(workbook,tabNames);
        List<String> rowL = VvpByQuarter.readDataL(workbook,tabNames);


        rowL.forEach(System.out::println);

        ConstructorsForXMLStatbase xml =  new ConstructorsForXMLStatbase();
        xml.ParamLangXML();
        xml.WriteParamXML(file.getName(),tabNames,rowA,rowB);
        writer.close();
    }
}
