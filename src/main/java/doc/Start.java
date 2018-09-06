package doc;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Start {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        List<String> list = new ArrayList<>();
        list = ReadDocuments.getFiles("/home/mdaniyar/Desktop/parser/link1");
        //get filenames from directory
        //System.out.println(list.get(0));
        //get paragraph name from document
        for (int i=0; i<list.size();i++) {
            System.out.println(ReadDocuments.getParagraph(list.get(i)));
            //First parameter домашних хозяйств
            String[] cells1=ReadDocuments.getCellName(list.get(i),"домашних хозяйств");
            System.out.println(cells1[0]+" "+cells1[1]);
            //Second parameter органов государственного управления
            String[] cells2=ReadDocuments.getCellName(list.get(i),"управления");
            if (cells2[1].isEmpty()){
                System.out.println("органов государственного "+cells2[0]+" "+cells2[2]);
            } else{
                System.out.println("органов государственного "+cells2[0]+" "+cells2[1]);
            }
            //Third parameter Валовое накопление
            String[] cells3=ReadDocuments.getCellName(list.get(i),"Валовое накопление");
            System.out.println(cells3[0]+" "+cells3[1]);

            //Forth parameter экспорт товаров и услуг
            String[] cells4=ReadDocuments.getCellName(list.get(i),"экспорт товаров и услуг");
            System.out.println(cells4[0]+" "+cells4[1]);

            //fifth parameter экспорт товаров и услуг
            String[] cells5=ReadDocuments.getCellName(list.get(i),"импорт товаров и услуг");
            System.out.println(cells5[0]+" "+cells5[1]);
            //Sith Валовой внутренний продукт методом конечного использования
            String[] cells6 = ReadDocuments.getCellName(list.get(i),"Валовой внутренний продукт методом производства");
            if (cells6[1].isEmpty()){
                System.out.println(cells6[0]+" "+cells6[2]);
            } else{
                System.out.println(cells6[0]+" "+cells6[1]);
            }
            System.out.println("===============================================================================================================");
        }




    }
}
