package doc;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTable;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ReadDocuments {

    public static  List<String> getFiles(String folderPath){
        List<String> list = new ArrayList<>();
        File folder = new File(folderPath);
        File[] listOfFiles = folder.listFiles();

        for (File file:
             listOfFiles) {
            if (file.isFile()){
                list.add(file.getAbsolutePath());
            }
        }
        return list;
    }

    public static String getParagraph(String filename) throws IOException, InvalidFormatException {
        FileInputStream fis  = new FileInputStream(new File(filename));
        XWPFDocument document = new XWPFDocument(fis);
        String paragraph = document.getParagraphs().get(2).getText();
        return  paragraph;
    }
    public static  String[] getCellName(String filename,String parameter) throws IOException, InvalidFormatException{
        FileInputStream fis  = new FileInputStream(new File(filename));
        XWPFDocument document = new XWPFDocument(fis);
        List<XWPFTable> table = document.getTables();
        String[] cell = null;
        String[] lines = table.get(1).getText().split("\\r?\\n");
        for (String  line:
             lines) {
            if (line.contains(parameter)){
                cell = line.split("\\t");
            }
        }
        return cell;
    }
}
