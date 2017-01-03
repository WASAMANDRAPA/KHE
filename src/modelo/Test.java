package modelo;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigInteger;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.jaxb.Context;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.wml.JcEnumeration;

public class Test {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

            File file = new File(System.getProperty("user.dir") + "/src/recursos/images/Consorcio_Vial_Jaylli.jpg");

            try (java.io.InputStream is = new java.io.FileInputStream(file)) {
                long length = file.length();
                if (length > Integer.MAX_VALUE) {
                    System.out.println("Imagen muy grande!");
                }
                byte[] bytes = new byte[(int) length];
                int offset = 0;
                int numRead = 0;
                while (offset < bytes.length && (numRead = is.read(bytes, offset, bytes.length - offset)) >= 0) {
                    offset += numRead;
                }
                if (offset < bytes.length) {
                    System.out.println("No se leyo el archivo completamente" + file.getName());
                }

                String filenameHint = null;
                String altText = null;
                int id1 = 0;
                int id2 = 1;

                org.docx4j.wml.P img = newImage(wordMLPackage, bytes, filenameHint, altText, id1, id2, 1, 4650);
                wordMLPackage.getMainDocumentPart().addObject(img);
            }

            MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
            titulo(wordMLPackage);
            mdp.addParagraphOfText("CHOCHOLOCO");

            String filename = System.getProperty("user.dir") + "/OUT_hello.docx";
            Docx4J.save(wordMLPackage, new java.io.File(filename), Docx4J.FLAG_SAVE_ZIP_FILE);
            System.out.println("Saved " + filename);

        } catch (InvalidFormatException ex) {
            System.out.println("Aviso: " + ex.getMessage());
        } catch (Docx4JException | FileNotFoundException ex) {
            System.out.println("Aviso: " + ex.getMessage());
        } catch (IOException ex) {
            System.out.println("Aviso: " + ex.getMessage());
        } catch (Exception ex) {
            System.out.println("Aviso: " + ex.getMessage());
        }
    }

    static org.docx4j.wml.P newImage(WordprocessingMLPackage wordMLPackage, byte[] bytes, String filenameHint, String altText, int id1, int id2, int posicion, long size) throws Exception {
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);
        Inline inline = imagePart.createImageInline(filenameHint, altText, id1, id2, size, false);
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.P p = factory.createP();
        org.docx4j.wml.R run = factory.createR();
        p.getContent().add(run);
        org.docx4j.wml.Drawing drawing = factory.createDrawing();

        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        p.setPPr(centrar(posicion));
        return p;
    }

    static org.docx4j.wml.P titulo(WordprocessingMLPackage wordMLPackage) {
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.P p = factory.createP();
        org.docx4j.wml.R run = factory.createR();
        org.docx4j.wml.Text text = factory.createText();
        //----------------------------------------------------------------------
        run.setRPr(formatoTitulo());
        //----------------------------------------------------------------------
        text.setValue("CERTIFICADO DE TRABAJO");

        run.getContent().add(text);
        p.getContent().add(run);

        p.setPPr(centrar(2));
        wordMLPackage.getMainDocumentPart().addObject(p);

        return p;
    }

    /**
     *
     * @param i (1 - Izquierda, 2 - centro, 3 - derecha)
     * @return FomartoTexto
     */
    static org.docx4j.wml.PPr centrar(int i) {
        org.docx4j.wml.PPr formatoTexto = Context.getWmlObjectFactory().createPPr();
        org.docx4j.wml.Jc justifitacion = Context.getWmlObjectFactory().createJc();

        switch (i) {
            case 1:
                justifitacion.setVal(JcEnumeration.LEFT);
                break;
            case 2:
                justifitacion.setVal(JcEnumeration.CENTER);
                break;
            case 3:
                justifitacion.setVal(JcEnumeration.RIGHT);
                break;
        }
        formatoTexto.setJc(justifitacion);
        return formatoTexto;
    }

    static org.docx4j.wml.RPr formatoTitulo() {

        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.RPr rPr = factory.createRPr();
        org.docx4j.wml.BooleanDefaultTrue b = new org.docx4j.wml.BooleanDefaultTrue();

        org.docx4j.wml.RFonts font = Context.getWmlObjectFactory().createRFonts();
        font.setAscii("Bookman Old Style");

        org.docx4j.wml.HpsMeasure sz = Context.getWmlObjectFactory().createHpsMeasure();
        sz.setVal(new BigInteger("44"));
        rPr.setRFonts(font);
        rPr.setSz(sz);
        b.setVal(true);
        rPr.setB(b);
        rPr.setI(b);

        return rPr;
    }
}
