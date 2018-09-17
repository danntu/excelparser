package html;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.net.MalformedURLException;
import java.util.ArrayList;

public class OilPrice1986Now {
    public static ArrayList<Elements> parseTable(String URL){
            ArrayList<Elements> data = new ArrayList<>();
        try{
            Document document = Jsoup.connect(URL).get();
            Element table = document.select("table").get(0);
            Elements rows = table.select("tr");
            data.add(rows.get(0).select("td"));// Select Table heading
            for (int j = 1; j<rows.size(); j++) {// Iterate through table data
                data.add(rows.get(j).select("td"));// Storing result in Array List
            }
        } catch (Exception e){
            System.out.println("Error by reading html file");
        }
        return data;
    }

    public static void main(String[] args) throws MalformedURLException {
        String URL = "https://worldtable.info/yekonomika/cena-na-neft-marki-brent-tablica-s-1986-po-20.html";
        ArrayList<Elements> tableData = parseTable(URL);
        int lastnumber = tableData.size()-1;
        int firstnumber = lastnumber-32;

        ArrayList<Elements> parsedTableData = new ArrayList<>();

        for (int i = firstnumber; i < lastnumber; i++) {
            parsedTableData.add(tableData.get(i));
        }

        parsedTableData.forEach(elements -> {
            //System.out.println(elements.text());
            System.out.println(elements.get(0).text()+" ; "+elements.get(1).text());
        });

    }
}
