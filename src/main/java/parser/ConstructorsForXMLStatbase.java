package parser;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ConstructorsForXMLStatbase {

    DocumentBuilder builder;

    public void ParamLangXML() {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        try {
            builder = factory.newDocumentBuilder();
        } catch (ParserConfigurationException e) {
            System.out.println("Error parsing to XML");
        }
    }

    public void WriteParamXML(String fileName, List<String> tabNames,
                              List<String> rowA, List<String> rowB) throws TransformerException, IOException{
        Document doc = builder.newDocument();
        Element RootElement1 = doc.createElement("root");
        Element RootElement2  = doc.createElement("file_name");
        RootElement2.appendChild(doc.createTextNode(fileName));
        Element RootElement3 = doc.createElement("tab");

        Element tabName;
        for (int i = 0; i < tabNames.size(); i++) {
            String tabNameForLoop = "tab"+(i+1);
            tabName = doc.createElement(tabNameForLoop);
            tabName.appendChild(doc.createTextNode(tabNames.get(i)));
            RootElement3.appendChild(tabName);
        }
        Element RootElement4 = doc.createElement("vvp1");
        RootElement4.appendChild(doc.createTextNode(rowA.get(0)));
        Element RootElement5 = doc.createElement("oked");
        RootElement5.appendChild(doc.createTextNode(rowA.get(1)));
        RootElement4.appendChild(RootElement5);
        Element RootElement6 = doc.createElement("year");
        RootElement6.appendChild(doc.createTextNode(rowB.get(0)));
        RootElement5.appendChild(RootElement6);

        Element RootElement7;
        Element RootElement8;
        for (int i = 2; i <=27 ; i++) {
            String element7 = "oked"+(i-1);
            String element8 = "oked"+(i-1)+"_price"+(i-1);
            RootElement7 = doc.createElement(element7);
            RootElement7.appendChild(doc.createTextNode(rowA.get(i)));
            RootElement6.appendChild(RootElement7);

            RootElement8 = doc.createElement(element8);
            RootElement8.appendChild(doc.createTextNode(rowB.get(i-1)));
            RootElement7.appendChild(RootElement8);
        }

        Element RootElement100 = doc.createElement("vvp2");
        RootElement100.appendChild(doc.createTextNode(rowA.get(28)));

        RootElement1.appendChild(RootElement2);
        RootElement1.appendChild(RootElement3);
        RootElement1.appendChild(RootElement4);
        RootElement1.appendChild(RootElement100);

        doc.appendChild(RootElement1);

        Transformer transformer = TransformerFactory.newInstance().newTransformer();
        transformer.transform(new DOMSource(doc), new StreamResult(new FileOutputStream("" +
                "/home/mdaniyar/mygit/excelparser/src/main/java/polygon/output/informationStatbase.xml")));

    }
}
