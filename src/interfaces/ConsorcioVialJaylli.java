package interfaces;

import java.awt.Component;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.xml.bind.JAXBException;
import org.docx4j.Docx4J;
import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Anchor;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.dml.wordprocessingDrawing.STAlignH;
import org.docx4j.dml.wordprocessingDrawing.STAlignV;
import org.docx4j.dml.wordprocessingDrawing.STRelFromH;
import org.docx4j.dml.wordprocessingDrawing.STRelFromV;
import org.docx4j.dml.wordprocessingDrawing.STWrapText;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.PPrBase;
import org.docx4j.wml.SectPr;

/**
 *
 * @author USUARIO
 */
public class ConsorcioVialJaylli extends javax.swing.JInternalFrame {

    String[] meses = {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"};
    static final Logger LOGGER = Logger.getLogger(Logger.GLOBAL_LOGGER_NAME);
    final JFileChooser fc = new JFileChooser();

    /**
     * Constructor por defecto
     */
    public ConsorcioVialJaylli() {
        initComponents();
        this.setResizable(false);
        addListeners();
        metodoTestSobreOcultarPaneles();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jbtnSalir = new javax.swing.JButton();
        jbtnCrear = new javax.swing.JButton();
        jPanel1 = new javax.swing.JPanel();
        jPanel2 = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        jlstDocumentos = new javax.swing.JList();
        jlblTest = new javax.swing.JLabel();
        jLayeredPane1 = new javax.swing.JLayeredPane();
        ConsorcioVialJaylli = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        jpanel1Nombre = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();
        jpanel1DNI = new javax.swing.JTextField();
        jpanel1FecInicio = new com.toedter.calendar.JDateChooser();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jpanel1FecFin = new com.toedter.calendar.JDateChooser();
        jLabel5 = new javax.swing.JLabel();
        jpanel1PoscLaboral = new javax.swing.JTextField();
        jpanelTest1 = new javax.swing.JPanel();
        jButton2 = new javax.swing.JButton();
        jpanelTest2 = new javax.swing.JPanel();
        jButton3 = new javax.swing.JButton();

        jbtnSalir.setFont(new java.awt.Font("Dialog", 1, 14)); // NOI18N
        jbtnSalir.setText("Salir");

        jbtnCrear.setFont(new java.awt.Font("Dialog", 1, 14)); // NOI18N
        jbtnCrear.setText("Crear");

        jPanel1.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));

        jPanel2.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));

        jlstDocumentos.setFont(new java.awt.Font("Dialog", 1, 12)); // NOI18N
        jlstDocumentos.setModel(new javax.swing.AbstractListModel() {
            String[] strings = { "Consorcio Via Jaylli", "A A", "B B", "C C" };
            public int getSize() { return strings.length; }
            public Object getElementAt(int i) { return strings[i]; }
        });
        jScrollPane1.setViewportView(jlstDocumentos);

        jlblTest.setText("TEST");

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 174, Short.MAX_VALUE)
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addComponent(jlblTest)
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 210, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 48, Short.MAX_VALUE)
                .addComponent(jlblTest)
                .addContainerGap())
        );

        jLayeredPane1.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));

        ConsorcioVialJaylli.setName("A"); // NOI18N

        jLabel1.setFont(new java.awt.Font("Dialog", 1, 12)); // NOI18N
        jLabel1.setText("Nombre Persona: ");

        jpanel1Nombre.setFont(new java.awt.Font("Dialog", 1, 12)); // NOI18N

        jLabel2.setFont(new java.awt.Font("Dialog", 1, 12)); // NOI18N
        jLabel2.setText("DNI Persona: ");

        jpanel1DNI.setFont(new java.awt.Font("Dialog", 1, 12)); // NOI18N

        jLabel3.setFont(new java.awt.Font("Dialog", 1, 12)); // NOI18N
        jLabel3.setText("Fecha Inicio Trabajo: ");

        jLabel4.setFont(new java.awt.Font("Dialog", 1, 12)); // NOI18N
        jLabel4.setText("Fecha Fin Trabajo: ");

        jLabel5.setFont(new java.awt.Font("Dialog", 1, 12)); // NOI18N
        jLabel5.setText("Posicion Laboral: ");

        javax.swing.GroupLayout ConsorcioVialJaylliLayout = new javax.swing.GroupLayout(ConsorcioVialJaylli);
        ConsorcioVialJaylli.setLayout(ConsorcioVialJaylliLayout);
        ConsorcioVialJaylliLayout.setHorizontalGroup(
            ConsorcioVialJaylliLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(ConsorcioVialJaylliLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(ConsorcioVialJaylliLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel1)
                    .addComponent(jLabel2)
                    .addComponent(jLabel3)
                    .addComponent(jLabel4)
                    .addComponent(jLabel5))
                .addGap(7, 7, 7)
                .addGroup(ConsorcioVialJaylliLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(ConsorcioVialJaylliLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(ConsorcioVialJaylliLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(ConsorcioVialJaylliLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                .addComponent(jpanel1DNI, javax.swing.GroupLayout.PREFERRED_SIZE, 200, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addComponent(jpanel1Nombre, javax.swing.GroupLayout.PREFERRED_SIZE, 200, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addComponent(jpanel1FecInicio, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 200, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addComponent(jpanel1FecFin, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 200, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(jpanel1PoscLaboral, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 200, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(167, Short.MAX_VALUE))
        );
        ConsorcioVialJaylliLayout.setVerticalGroup(
            ConsorcioVialJaylliLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(ConsorcioVialJaylliLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(ConsorcioVialJaylliLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel1)
                    .addComponent(jpanel1Nombre, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(ConsorcioVialJaylliLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(jpanel1DNI, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(15, 15, 15)
                .addGroup(ConsorcioVialJaylliLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel3)
                    .addComponent(jpanel1FecInicio, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(ConsorcioVialJaylliLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel4)
                    .addComponent(jpanel1FecFin, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(ConsorcioVialJaylliLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel5)
                    .addComponent(jpanel1PoscLaboral, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(137, Short.MAX_VALUE))
        );

        jpanelTest1.setName("B"); // NOI18N

        jButton2.setText("jButton2");

        javax.swing.GroupLayout jpanelTest1Layout = new javax.swing.GroupLayout(jpanelTest1);
        jpanelTest1.setLayout(jpanelTest1Layout);
        jpanelTest1Layout.setHorizontalGroup(
            jpanelTest1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jpanelTest1Layout.createSequentialGroup()
                .addGap(204, 204, 204)
                .addComponent(jButton2)
                .addContainerGap(228, Short.MAX_VALUE))
        );
        jpanelTest1Layout.setVerticalGroup(
            jpanelTest1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jpanelTest1Layout.createSequentialGroup()
                .addContainerGap(137, Short.MAX_VALUE)
                .addComponent(jButton2)
                .addGap(135, 135, 135))
        );

        jpanelTest2.setName("C"); // NOI18N

        jButton3.setText("jButton3");

        javax.swing.GroupLayout jpanelTest2Layout = new javax.swing.GroupLayout(jpanelTest2);
        jpanelTest2.setLayout(jpanelTest2Layout);
        jpanelTest2Layout.setHorizontalGroup(
            jpanelTest2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jpanelTest2Layout.createSequentialGroup()
                .addGap(207, 207, 207)
                .addComponent(jButton3)
                .addContainerGap(225, Short.MAX_VALUE))
        );
        jpanelTest2Layout.setVerticalGroup(
            jpanelTest2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jpanelTest2Layout.createSequentialGroup()
                .addContainerGap(242, Short.MAX_VALUE)
                .addComponent(jButton3)
                .addGap(30, 30, 30))
        );

        javax.swing.GroupLayout jLayeredPane1Layout = new javax.swing.GroupLayout(jLayeredPane1);
        jLayeredPane1.setLayout(jLayeredPane1Layout);
        jLayeredPane1Layout.setHorizontalGroup(
            jLayeredPane1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(ConsorcioVialJaylli, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(jLayeredPane1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addComponent(jpanelTest1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
            .addGroup(jLayeredPane1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addComponent(jpanelTest2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jLayeredPane1Layout.setVerticalGroup(
            jLayeredPane1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(ConsorcioVialJaylli, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(jLayeredPane1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addComponent(jpanelTest1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
            .addGroup(jLayeredPane1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addComponent(jpanelTest2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jLayeredPane1.setLayer(ConsorcioVialJaylli, javax.swing.JLayeredPane.DEFAULT_LAYER);
        jLayeredPane1.setLayer(jpanelTest1, javax.swing.JLayeredPane.DEFAULT_LAYER);
        jLayeredPane1.setLayer(jpanelTest2, javax.swing.JLayeredPane.DEFAULT_LAYER);

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLayeredPane1))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addComponent(jLayeredPane1)
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(22, 22, 22)
                .addComponent(jbtnCrear, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 502, Short.MAX_VALUE)
                .addComponent(jbtnSalir, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jbtnSalir)
                    .addComponent(jbtnCrear))
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    /**
     * Procedimiento que agrega los diferentes Listeners a los componentes
     * dentro del aplicativo
     */
    private void addListeners() {
        /**
         * Evento del boton Salir
         */
        jbtnSalir.addActionListener((ActionEvent ae) -> {
            this.dispose();
        });
        /**
         * Evento del boton Crear
         */
        jbtnCrear.addActionListener((ActionEvent ae) -> {
            metodoGuardar();
        });
        /**
         * Evento doble click de la Lista de documentos
         */
        jlstDocumentos.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent evt) {
                JList list = (JList) evt.getSource();
                if (evt.getClickCount() == 2) {
                    if (evt.getButton() == MouseEvent.BUTTON1) {
                        Rectangle r = list.getCellBounds(0, list.getLastVisibleIndex());
                        if (r != null && r.contains(evt.getPoint())) {
                            int index = list.locationToIndex(evt.getPoint());
                            iniciarPanel(index);
                        }
                    }
                }
            }
        });
    }

    private void metodoGuardar() {
        int res = JOptionPane.showConfirmDialog(this, "Desea especificar la ruta de guardado", "Aviso del Sistema", JOptionPane.YES_NO_CANCEL_OPTION);
        switch (res) {
            case JOptionPane.YES_OPTION:
                int respuesta = fc.showSaveDialog(this);
                if (respuesta == JFileChooser.APPROVE_OPTION) {
                    File file = fc.getSelectedFile();
                    //LOGGER.info(file.toString());
                    crearDocxFile(file.getAbsolutePath());
                    JOptionPane.showMessageDialog(this, "Archivo Guardado", "Aviso del Sistema", JOptionPane.INFORMATION_MESSAGE);
                }
                break;
            case JOptionPane.NO_OPTION:
                crearDocxFile("");
                JOptionPane.showMessageDialog(this, "Archivo Guardado", "Aviso del Sistema", JOptionPane.INFORMATION_MESSAGE);
                break;
            case JOptionPane.CANCEL_OPTION:
                break;
        }
    }

    private void iniciarPanel(int index) {
        ConsorcioVialJaylli.setVisible(true);
        jlblTest.setText(String.valueOf(index));
    }

    private void metodoTestSobreOcultarPaneles() {
        Component[] cmpts = jLayeredPane1.getComponents();
        for (Component cmpt : cmpts) {
            System.out.println(cmpt.getName());
            cmpt.setVisible(false);
        }
    }

    private void crearDocxFile(String ruta) {

        try {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage(org.docx4j.model.structure.PageSizePaper.A4, false);

            // -----------------------------------------------------------------
            //wordMLPackage.getMainDocumentPart().getJaxbElement().getBody().setSectPr(testingSangria());
            // -----------------------------------------------------------------
            agregarPrimeraImagen(wordMLPackage);
            // -----------------------------------------------------------------
            wordMLPackage.getMainDocumentPart().addObject(titulo());
            // -----------------------------------------------------------------
            wordMLPackage.getMainDocumentPart().addParagraphOfText("");
            // -----------------------------------------------------------------
            wordMLPackage.getMainDocumentPart().addObject(primerCuerpo());
            // -----------------------------------------------------------------
            wordMLPackage.getMainDocumentPart().addObject(segundoCuerpo());
            // -----------------------------------------------------------------
            wordMLPackage.getMainDocumentPart().addObject(tercerCuerpo());
            // -----------------------------------------------------------------
            wordMLPackage.getMainDocumentPart().addObject(cuartoCuerpo());
            // -----------------------------------------------------------------
            wordMLPackage.getMainDocumentPart().addObject(parteDesdeXml1());
            // -----------------------------------------------------------------
            agregarUltimaImagen(wordMLPackage);
            // -----------------------------------------------------------------
            System.out.println("ACA");
            if (ruta.trim().length() <= 0) {
                String filename = System.getProperty("user.dir") + "/OUT_hello.docx";
                Docx4J.save(wordMLPackage, new java.io.File(filename), Docx4J.FLAG_SAVE_ZIP_FILE);
            } else {
                Docx4J.save(wordMLPackage, new java.io.File(ruta), Docx4J.FLAG_SAVE_ZIP_FILE);
            }

            System.out.println("Saved");

        } catch (InvalidFormatException ex) {
            System.out.println("Aviso InvalidFormatException: " + ex.getMessage());
        } catch (Docx4JException | FileNotFoundException ex) {
            System.out.println("Aviso Docx4JException o FileNotFoundException: " + ex.getMessage());
        } catch (IOException ex) {
            System.out.println("Aviso IOException: " + ex.getMessage());
        } catch (Exception ex) {
            System.out.println("Aviso Exception: " + ex.getMessage());
        }
    }

    private org.docx4j.wml.P primerCuerpo() {
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.P p = factory.createP();

        org.docx4j.wml.R run = factory.createR();
        org.docx4j.wml.Text text = factory.createText();
        text.setSpace("preserve");
        text.setValue("CONSORCIO  VIAL  JAYLLI,   ");
        run.getContent().add(text);
        run.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), true, true));

        org.docx4j.wml.R run2 = factory.createR();
        org.docx4j.wml.Text text2 = factory.createText();
        text.setSpace("preserve");
        text2.setValue("con  RUC Nº 20566464773,    con     dirección    en   Av.  Nicolás Ayllon 2634 - Ate - Lima, certifica: ");
        run2.getContent().add(text2);
        run2.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), false, true));

        p.getContent().add(run);
        p.getContent().add(run2);

        p.setPPr(posicion(4, false));

        return p;
    }

    private org.docx4j.wml.P segundoCuerpo() {
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.P p = factory.createP();

        org.docx4j.wml.R run = factory.createR();
        org.docx4j.wml.Text text = factory.createText();
        text.setSpace("preserve");
        text.setValue("Que el/la Señor(a), ");
        run.getContent().add(text);
        run.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), false, true));

        org.docx4j.wml.R run2 = factory.createR();
        org.docx4j.wml.Text text2 = factory.createText();
        text.setSpace("preserve");
        text2.setValue(jpanel1Nombre.getText().trim());
        run2.getContent().add(text2);
        run2.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), true, true));

        org.docx4j.wml.R run3 = factory.createR();
        org.docx4j.wml.Text text3 = factory.createText();
        text3.setSpace("preserve");
        text3.setValue(" con ");
        run3.getContent().add(text3);
        run3.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), false, true));

        org.docx4j.wml.R run4 = factory.createR();
        org.docx4j.wml.Text text4 = factory.createText();
        text4.setSpace("preserve");
        text4.setValue("DNI " + jpanel1DNI.getText().trim() + ", ");
        run4.getContent().add(text4);
        run4.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), true, true));

        org.docx4j.wml.R run5 = factory.createR();
        org.docx4j.wml.Text text5 = factory.createText();
        text5.setSpace("preserve");
        text5.setValue("laboro en nuestra empresa desde el ");
        run5.getContent().add(text5);
        run5.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), false, true));

        org.docx4j.wml.R run6 = factory.createR();
        org.docx4j.wml.Text text6 = factory.createText();
        text6.setSpace("preserve");
        text6.setValue(asignarFormato(jpanel1FecInicio.getDate()));
        run6.getContent().add(text6);
        run6.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), false, true));

        org.docx4j.wml.R run7 = factory.createR();
        org.docx4j.wml.Text text7 = factory.createText();
        text7.setSpace("preserve");
        text7.setValue(" hasta el ");
        run7.getContent().add(text7);
        run7.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), false, true));

        org.docx4j.wml.R run8 = factory.createR();
        org.docx4j.wml.Text text8 = factory.createText();
        text8.setSpace("preserve");
        text8.setValue(asignarFormato(jpanel1FecFin.getDate()));
        run8.getContent().add(text8);
        run8.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), false, true));

        org.docx4j.wml.R run9 = factory.createR();
        org.docx4j.wml.Text text9 = factory.createText();
        text9.setSpace("preserve");
        text9.setValue(", desempeñando el cargo de ");
        run9.getContent().add(text9);
        run9.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), false, true));

        org.docx4j.wml.R run10 = factory.createR();
        org.docx4j.wml.Text text10 = factory.createText();
        text10.setSpace("preserve");
        text10.setValue(jpanel1PoscLaboral.getText().trim());
        run10.getContent().add(text10);
        run10.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), true, true));

        org.docx4j.wml.R run11 = factory.createR();
        org.docx4j.wml.Text text11 = factory.createText();
        text11.setSpace("preserve");
        text11.setValue(" en la obra ");
        run11.getContent().add(text11);
        run11.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), false, true));

        org.docx4j.wml.R run12 = factory.createR();
        org.docx4j.wml.Text text12 = factory.createText();
        text12.setSpace("preserve");
        text12.setValue(" \"Rehabilitación y Mejoramiento de la Carretera Huancavelica Lircay, tramo: Km 1+550 (Av. Los Chancas) - Lircay\", ");
        run12.getContent().add(text12);
        run12.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), true, true));

        org.docx4j.wml.R run13 = factory.createR();
        org.docx4j.wml.Text text13 = factory.createText();
        text13.setSpace("preserve");
        text13.setValue(" la cual está ubicada en el Distrito de Huancavelica y Lircay, Provincias de Huancavelica y Angaraes, Departamento de Huancavelica.");
        run13.getContent().add(text13);
        run13.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), false, true));

        p.getContent().add(run);
        p.getContent().add(run2);
        p.getContent().add(run3);
        p.getContent().add(run4);
        p.getContent().add(run5);
        p.getContent().add(run6);
        p.getContent().add(run7);
        p.getContent().add(run8);
        p.getContent().add(run9);
        p.getContent().add(run10);
        p.getContent().add(run11);
        p.getContent().add(run12);
        p.getContent().add(run13);

        p.setPPr(posicion(4, false));

        return p;
    }

    private org.docx4j.wml.P tercerCuerpo() {
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.P p = factory.createP();

        org.docx4j.wml.R run = factory.createR();
        org.docx4j.wml.Text text = factory.createText();
        text.setSpace("preserve");
        text.setValue("Se expide el presente certificado, a solicitud del interesado para los fines que estime conveniente.");
        run.getContent().add(text);
        run.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), false, true));

        p.getContent().add(run);

        p.setPPr(posicion(4, false));

        return p;
    }

    private org.docx4j.wml.P cuartoCuerpo() {
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.P p = factory.createP();

        org.docx4j.wml.R run = factory.createR();
        org.docx4j.wml.Text text = factory.createText();
        text.setSpace("preserve");
        text.setValue("Huancavelica, 27 de Febrero del 2016");
        run.getContent().add(text);
        run.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("32"), false, true));

        p.getContent().add(run);

        p.setPPr(posicion(2, true));

        return p;
    }

    private Object parteDesdeXml1() throws JAXBException {
        String xml = "<w:p  xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "                    <w:pPr>\n"
                + "                        <w:pStyle w:val=\"Style3\"/>\n"
                + "                        <w:framePr w:w=\"4000\" w:h=\"500\" w:hSpace=\"38\" w:wrap=\"notBeside\" w:hAnchor=\"page\" w:vAnchor=\"text\" w:x=\"2000\" w:y=\"1909\" w:hRule=\"exact\"/>\n"
                + "                        <w:widowControl/>\n"
                + "                        <w:rPr>\n"
                + "                            <w:rStyle w:val=\"FontStyle14\"/>\n"
                + "                            <w:sz w:val=\"22\"/>\n"
                + "                            <w:lang w:val=\"es-ES_tradnl\" w:eastAsia=\"es-ES_tradnl\"/>\n"
                + "                        </w:rPr>\n"
                + "                    </w:pPr>\n"
                + "                    <w:r>\n"
                + "                        <w:rPr>\n"
                + "                            <w:rStyle w:val=\"FontStyle14\"/>\n"
                + "                            <w:sz w:val=\"22\"/>\n"
                + "                            <w:lang w:val=\"es-ES_tradnl\" w:eastAsia=\"es-ES_tradnl\"/>\n"
                + "                        </w:rPr>\n"
                + "                        <w:t>CONSORCIO VIAL JAYLLI</w:t>\n"
                + "                    </w:r>\n"
                + "                    <w:r>\n"
                + "                    <w:br/>\n"
                + "                    </w:r>\n"
                + "                    <w:r>\n"
                + "                        <w:rPr>\n"
                + "                            <w:rStyle w:val=\"FontStyle14\"/>\n"
                + "                            <w:sz w:val=\"22\"/>\n"
                + "                            <w:lang w:val=\"es-ES_tradnl\" w:eastAsia=\"es-ES_tradnl\"/>\n"
                + "                        </w:rPr>\n"
                + "                        <w:t>Av. Nicolás Ayllón Nº 2634-Ate-Lima 3-Perú</w:t>\n"
                + "                    </w:r>\n"
                + "                </w:p>";
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.P pepa = factory.createP();
        pepa = (org.docx4j.wml.P) org.docx4j.XmlUtils.unmarshalString(xml);
        //pepa.getContent().add(org.docx4j.XmlUtils.unmarshalString(xml));

        return pepa;
    }

    private org.docx4j.wml.SectPr testingSangria() {
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        SectPr sectPr = factory.createSectPr();
        SectPr.PgMar pgMar = new SectPr.PgMar();
        pgMar.setTop(new BigInteger("1417"));
        pgMar.setBottom(new BigInteger("1417"));
        pgMar.setRight(new BigInteger("1700"));
        pgMar.setLeft(new BigInteger("1701"));
        pgMar.setHeader(new BigInteger("708"));
        pgMar.setFooter(new BigInteger("708"));
        pgMar.setGutter(new BigInteger("0"));
        sectPr.setPgMar(pgMar);

        return sectPr;
    }

    private void agregarPrimeraImagen(WordprocessingMLPackage wordMLPackage) throws Exception {
        File image1 = new File(System.getProperty("user.dir") + "/src/recursos/images/Consorcio_Vial_Jaylli.jpg");
        org.docx4j.wml.P img1 = newImage(image1, wordMLPackage, 4650);
        img1.setPPr(posicion(1, false));
        wordMLPackage.getMainDocumentPart().addObject(img1);
    }

    private void agregarUltimaImagen(WordprocessingMLPackage wordMLPackage) throws Exception {
        File image2 = new File(System.getProperty("user.dir") + "/src/recursos/images/Firma_Consorcio_Vial_Jaylli.jpg");
        org.docx4j.wml.P img2 = newImage2(image2, wordMLPackage, 3900);
        img2.setPPr(posicion(3, true));
        wordMLPackage.getMainDocumentPart().addObject(img2);
    }

    private org.docx4j.wml.P newImage(File file, WordprocessingMLPackage wordMLPackage, long size)
            throws Exception, IOException, FileNotFoundException, Docx4JException {
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

            BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);
            Inline inline = imagePart.createImageInline(filenameHint, altText, id1, id2, size, false);
            org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
            org.docx4j.wml.P p = factory.createP();
            org.docx4j.wml.R run = factory.createR();
            org.docx4j.wml.Drawing drawing = factory.createDrawing();   
            
            drawing.getAnchorOrInline().add(inline);
            run.getContent().add(drawing);
            p.getContent().add(run);
            return p;
        }
    }
    
    private org.docx4j.wml.P newImage2(File file, WordprocessingMLPackage wordMLPackage, long size)
            throws Exception, IOException, FileNotFoundException, Docx4JException {
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

            BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);
            Inline inline = imagePart.createImageInline(filenameHint, altText, id1, id2, size, false);
            org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
            org.docx4j.wml.P p = factory.createP();
            org.docx4j.wml.R run = factory.createR();
            org.docx4j.wml.Drawing drawing = factory.createDrawing();
            
            // transforma inline to anchor
            String anchorXml = XmlUtils.marshaltoString(inline, true, false, Context.jc, Namespaces.NS_WORD12, "anchor", Inline.class);
            
            org.docx4j.dml.ObjectFactory dmlFactory = new org.docx4j.dml.ObjectFactory();
            org.docx4j.dml.wordprocessingDrawing.ObjectFactory wordDmlFactory = new org.docx4j.dml.wordprocessingDrawing.ObjectFactory();
            
            Anchor anchor = (Anchor)XmlUtils.unmarshalString(anchorXml, Context.jc, Anchor.class);
            anchor.setSimplePos(dmlFactory.createCTPoint2D());
            anchor.getSimplePos().setX(0L);
            anchor.getSimplePos().setY(0L);
            anchor.setSimplePosAttr(false);
            anchor.setPositionH(wordDmlFactory.createCTPosH());
            anchor.getPositionH().setAlign(STAlignH.RIGHT);
            anchor.getPositionH().setRelativeFrom(STRelFromH.MARGIN);
            anchor.setPositionV(wordDmlFactory.createCTPosV());
            anchor.getPositionV().setAlign(STAlignV.BOTTOM);
            anchor.getPositionV().setRelativeFrom(STRelFromV.MARGIN);
            // -----------------------------------------------------------------
            //anchor.setWrapNone(wordDmlFactory.createCTWrapNone());
            // -----------------------------------------------------------------
            anchor.setWrapSquare(wordDmlFactory.createCTWrapSquare());
            anchor.getWrapSquare().setWrapText(STWrapText.BOTH_SIDES);
            
            drawing.getAnchorOrInline().add(anchor);
            run.getContent().add(drawing);
            p.getContent().add(run);
            return p;
        }
    }

    private org.docx4j.wml.P titulo() {
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.P p = factory.createP();

        org.docx4j.wml.R run = factory.createR();
        org.docx4j.wml.Text text = factory.createText();
        text.setValue("CERTIFICADO DE TRABAJO");
        run.getContent().add(text);
        run.setRPr(fuenteTipoTamañoNegritaCursiva("Bookman Old Style", new BigInteger("44"), true, true));

        p.getContent().add(run);

        p.setPPr(posicion(2, false));

        return p;
    }

    private org.docx4j.wml.PPr posicion(int i, boolean espacio) {
        org.docx4j.wml.PPr formatoTexto = Context.getWmlObjectFactory().createPPr();
        org.docx4j.wml.Jc justifitacion = Context.getWmlObjectFactory().createJc();

        if (espacio){
            org.docx4j.wml.PPrBase.Spacing space = new PPrBase.Spacing();
            space.setAfter(BigInteger.ZERO);
            space.setBefore(BigInteger.ZERO);
            formatoTexto.setSpacing(space);
        }
        
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
            case 4:
                justifitacion.setVal(JcEnumeration.BOTH);
                break;
        }
        formatoTexto.setJc(justifitacion);
        return formatoTexto;
    }

    private org.docx4j.wml.RPr fuenteTipoTamañoNegritaCursiva(String tipo, BigInteger tamaño, boolean negrita, boolean cursiva) {
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.RPr rPr = factory.createRPr();
        org.docx4j.wml.RFonts font = Context.getWmlObjectFactory().createRFonts();
        org.docx4j.wml.HpsMeasure sz = Context.getWmlObjectFactory().createHpsMeasure();
        org.docx4j.wml.BooleanDefaultTrue b = new org.docx4j.wml.BooleanDefaultTrue();
        b.setVal(true);
        font.setAscii(tipo);
        sz.setVal(tamaño);
        if (negrita) {
            rPr.setB(b);
        }
        if (cursiva) {
            rPr.setI(b);
        }
        rPr.setSz(sz);
        rPr.setRFonts(font);

        return rPr;
    }

    private String asignarFormato(Date date) {
        String fec = null;
        if (date != null) {
            DateFormat formato = new SimpleDateFormat("dd/MM/yyyy");
            fec = formato.format(date);
            String[] partes = fec.split("/");
            partes[1] = meses[Integer.valueOf(partes[1]) - 1];
            fec = partes[0] + " de " + partes[1] + " del " + partes[2];
            //System.out.println(fec);
        }
        return fec;
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JPanel ConsorcioVialJaylli;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLayeredPane jLayeredPane1;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JButton jbtnCrear;
    private javax.swing.JButton jbtnSalir;
    private javax.swing.JLabel jlblTest;
    private javax.swing.JList jlstDocumentos;
    private javax.swing.JTextField jpanel1DNI;
    private com.toedter.calendar.JDateChooser jpanel1FecFin;
    private com.toedter.calendar.JDateChooser jpanel1FecInicio;
    private javax.swing.JTextField jpanel1Nombre;
    private javax.swing.JTextField jpanel1PoscLaboral;
    private javax.swing.JPanel jpanelTest1;
    private javax.swing.JPanel jpanelTest2;
    // End of variables declaration//GEN-END:variables
}
