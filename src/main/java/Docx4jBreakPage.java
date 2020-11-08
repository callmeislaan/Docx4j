import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Br;
import org.docx4j.wml.STBrType;

import java.io.File;

public class Docx4jBreakPage {

    public static void copyTemplate(MainDocumentPart mainDocumentPart, MainDocumentPart templateMainDocumentPart) {
        for (Object object :
                templateMainDocumentPart.getContent()) {
            mainDocumentPart.addObject(object);
        }
    }

    public static void breakPage(MainDocumentPart mainDocumentPart, MainDocumentPart templateMainDocumentPart) {
        Br pageBreak = Context.getWmlObjectFactory().createBr();
        pageBreak.setType(STBrType.PAGE);
        mainDocumentPart.addObject(pageBreak);
    }

    public static void main(String[] args) throws Docx4JException {
        WordprocessingMLPackage template = WordprocessingMLPackage.load(new File("D:\\Java\\Docx4j\\template.docx"));
        MainDocumentPart templateMainDocumentPart = template.getMainDocumentPart();

        WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
        MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();

        copyTemplate(mainDocumentPart, templateMainDocumentPart);

        breakPage(mainDocumentPart, templateMainDocumentPart);

        copyTemplate(mainDocumentPart, templateMainDocumentPart);

        wordPackage.save(new File("D:\\Java\\Docx4j\\copy_template.docx"));
    }
}
