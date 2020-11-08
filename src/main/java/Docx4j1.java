import com.rits.cloning.Cloner;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class Docx4j1 {

    public static void copyTemplate(MainDocumentPart mainDocumentPart, MainDocumentPart templateMainDocumentPart) {
        Cloner cloner=new Cloner();
        for (Object object :
                templateMainDocumentPart.getContent()) {
            mainDocumentPart.addObject(cloner.deepClone(object));
        }
    }

    public static void breakPage(MainDocumentPart mainDocumentPart) {
        Br pageBreak = Context.getWmlObjectFactory().createBr();
        pageBreak.setType(STBrType.PAGE);
        mainDocumentPart.addObject(pageBreak);
    }


    public static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        if (obj.getClass().equals(toSearch))
            result.add(obj);
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }

    private static void replaceTable(String[] placeholders, List<Map<String, String>> textToAdd, MainDocumentPart mainDocumentPart) throws Docx4JException, JAXBException {
        List<Object> tables = getAllElementFromObject(mainDocumentPart, Tbl.class);
        Tbl tempTable = getTemplateTable(tables, placeholders[0]);
        List<Object> rows = getAllElementFromObject(tempTable, Tr.class);
//        if (rows.size() == 1) { //careful only tables with 1 row are considered here
            Tr templateRow = (Tr) rows.get(rows.size()-1);
            for (Map<String, String> replacements : textToAdd) {
                addRowToTable(tempTable, templateRow, replacements);
            }
            assert tempTable != null;
            tempTable.getContent().remove(templateRow);
//        }
    }

    private static void addRowToTable(Tbl reviewTable, Tr templateRow, Map<String, String> replacements) {
        Tr workingRow = XmlUtils.deepCopy(templateRow);
        List<?> textElements = getAllElementFromObject(workingRow, Text.class);
        for (Object object : textElements) {
            Text text = (Text) object;
            String replacementValue = replacements.get(text.getValue());
            if (replacementValue != null)
                text.setValue(replacementValue);
        }
        reviewTable.getContent().add(workingRow);
    }

    private static Tbl getTemplateTable(List<Object> tables, String templateKey) {
        for (Object tbl : tables) {
            List<?> textElements = getAllElementFromObject(tbl, Text.class);
            for (Object text : textElements) {
                Text textElement = (Text) text;
                if (textElement.getValue() != null && textElement.getValue().equals(templateKey))
                    return (Tbl) tbl;
            }
        }
        return null;
    }

    public static void main(String[] args) throws Exception {
        WordprocessingMLPackage template = WordprocessingMLPackage.load(new File("D:\\Java\\Docx4j\\template.docx"));
        MainDocumentPart templateMainDocumentPart = template.getMainDocumentPart();

        WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
        MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();

        copyTemplate(mainDocumentPart, templateMainDocumentPart);


        List<Map<String, String>> list = new ArrayList<>();

        Map<String, String> entry = new HashMap<>();

        entry.put("${no}", "1");
        entry.put("${orderName}", "socola");
        entry.put("${quantity}", "56");
        list.add(entry);

        entry = new HashMap<>();
        entry.put("${no}", "2");
        entry.put("${orderName}", "milk");
        entry.put("${quantity}", "6");
        list.add(entry);

        replaceTable(new String[]{"${no}"}, list, mainDocumentPart);

        Map<String, String> mappings = new HashMap<>();

        VariablePrepare.prepare(wordPackage);

        mappings.put("total", "1234");
        mappings.put("name", "phuoc");
        mappings.put("address", "ha noi");

        wordPackage.getMainDocumentPart().variableReplace(mappings);

        ////////////////////////////////////////////////////////////////////////////////////////

        breakPage(mainDocumentPart);
        copyTemplate(mainDocumentPart, templateMainDocumentPart);

        List<Map<String, String>> list1 = new ArrayList<>();

        Map<String, String> entry1 = new HashMap<>();

        entry1.put("${no}", "3");
        entry1.put("${orderName}", "pate");
        entry1.put("${quantity}", "10");
        list1.add(entry1);

        entry1 = new HashMap<>();
        entry1.put("${no}", "4");
        entry1.put("${orderName}", "bread");
        entry1.put("${quantity}", "25");
        list1.add(entry1);

        entry1 = new HashMap<>();
        entry1.put("${no}", "5");
        entry1.put("${orderName}", " pizza");
        entry1.put("${quantity}", "16");
        list1.add(entry1);

        replaceTable(new String[]{"${no}"}, list1, mainDocumentPart);


        mappings = new HashMap<>();

        VariablePrepare.prepare(wordPackage);
        mappings.put("total", "634");
        mappings.put("name", "hoang");
        mappings.put("address", "nam dinh");
        wordPackage.getMainDocumentPart().variableReplace(mappings);


        wordPackage.save(new File("D:\\Java\\Docx4j\\result_1234.docx"));
    }

}
