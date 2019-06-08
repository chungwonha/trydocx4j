import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.List;

public class Main {

    public static void main(String[]  s){
        Main mymain = new Main();
        mymain.run();
    }

    public void run(){

        try {
            WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
            MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();
            mainDocumentPart.addStyledParagraphOfText("Title", "Hello World!");
            mainDocumentPart.addParagraphOfText("HELLO WORLD!!!!");
            P p = this.stylize();

            File image = new File("panamera.jpg" );
            byte[] fileContent = Files.readAllBytes(image.toPath());
            BinaryPartAbstractImage imagePart = BinaryPartAbstractImage
                    .createImagePart(wordPackage, fileContent);
            Inline inline = imagePart.createImageInline(
                    "Baeldung Image (filename hint)", "Alt Text", 1, 2, false);
            P Imageparagraph = addImageToParagraph(inline);

            int writableWidthTwips = wordPackage.getDocumentModel()
                    .getSections().get(0).getPageDimensions().getWritableWidthTwips();
            int columnNumber = 3;
            Tbl tbl = TblFactory.createTable(3, 3, writableWidthTwips/columnNumber);
            List<Object> rows = tbl.getContent();
            for (Object row : rows) {
                Tr tr = (Tr) row;
                List<Object> cells = tr.getContent();
                for(Object cell : cells) {
                    Tc td = (Tc) cell;
                    td.getContent().add(Imageparagraph);
                }
            }

            mainDocumentPart.getContent().add(tbl);
            mainDocumentPart.getContent().add(Imageparagraph);

            File exportFile = new File("welcome.docx");
            wordPackage.save(exportFile);
        }catch(IOException e) {
            e.printStackTrace();
        }catch(Docx4JException e) {
            e.printStackTrace();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public P stylize(){
        ObjectFactory factory = Context.getWmlObjectFactory();
        P p = factory.createP();
        R r = factory.createR();
        Text t = factory.createText();
        t.setValue("Welcome To Baeldung");
        r.getContent().add(t);
        p.getContent().add(r);
        RPr rpr = factory.createRPr();
        BooleanDefaultTrue b = new BooleanDefaultTrue();
        rpr.setB(b);
        rpr.setI(b);
        rpr.setCaps(b);
        Color green = factory.createColor();
        green.setVal("green");
        rpr.setColor(green);
        r.setRPr(rpr);
        return p;
    }

    private P addImageToParagraph(Inline inline) {



        ObjectFactory factory = new ObjectFactory();
        P p = factory.createP();
        R r = factory.createR();
        p.getContent().add(r);
        Drawing drawing = factory.createDrawing();
        r.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        return p;
    }

    public void addTable(WordprocessingMLPackage wordPackage, P p){

        int writableWidthTwips = wordPackage.getDocumentModel()
                .getSections().get(0).getPageDimensions().getWritableWidthTwips();
        int columnNumber = 3;
        Tbl tbl = TblFactory.createTable(3, 3, writableWidthTwips/columnNumber);
        List<Object> rows = tbl.getContent();
        for (Object row : rows) {
            Tr tr = (Tr) row;
            List<Object> cells = tr.getContent();
            for(Object cell : cells) {
                Tc td = (Tc) cell;
                td.getContent().add(p);
            }
        }

    }
}
