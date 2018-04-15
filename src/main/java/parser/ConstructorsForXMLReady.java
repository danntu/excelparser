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

public class ConstructorsForXMLReady {

    DocumentBuilder builder;

    public void ParamLangXML() {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        try {
            builder = factory.newDocumentBuilder();
        } catch (ParserConfigurationException e) {
            System.out.println("Error parsing to XML");
        }
    }

    public void WriteParamXML(List<String> tabNames) throws TransformerException, IOException{
        Document doc = builder.newDocument();
        Element RootElement1 = doc.createElement("root");
        Element RootElement2 = doc.createElement("tab");

        Element tabName;
        for (int i = 0; i < tabNames.size(); i++) {
            String tabNameForLoop = "tab"+(i+1);
            tabName = doc.createElement(tabNameForLoop);
            tabName.appendChild(doc.createTextNode(tabNames.get(i)));
            RootElement2.appendChild(tabName);
        }
        RootElement1.appendChild(RootElement2);
        doc.appendChild(RootElement1);

        Transformer transformer = TransformerFactory.newInstance().newTransformer();
        transformer.transform(new DOMSource(doc), new StreamResult(new FileOutputStream("" +
                "/home/mdaniyar/mygit/excelparser/src/main/java/polygon/output/informationReady.xml")));

    }
}
