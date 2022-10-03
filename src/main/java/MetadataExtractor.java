import org.apache.jempbox.xmp.XMPMetadata;
import org.apache.jempbox.xmp.XMPSchemaBasic;
import org.apache.jempbox.xmp.XMPSchemaDublinCore;
import org.apache.jempbox.xmp.XMPSchemaPDF;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentCatalog;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.common.PDMetadata;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.interactive.action.PDActionURI;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotationLink;

import netscape.javascript.JSException;

import javax.swing.*;
import javax.swing.plaf.ColorChooserUI;
import javax.swing.text.AttributeSet.FontAttribute;

import java.awt.*;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.List;
import java.util.Objects;
import java.util.TimeZone;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Date;

public class MetadataExtractor {

    private static final JTextArea journal = new JTextArea(); // where messages are output
    private static final JScrollPane sp = new JScrollPane(journal);
    private static final JProgressBar progress = new JProgressBar();

    private static String[] PDF_List = new String[0];

    public static void main(String[] args) {

        JFrame sfc = new JFrame("PDF Metadata Extractor");

        sfc.setSize(800, 600);
        sfc.setPreferredSize(new Dimension(800, 600));
        sfc.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        sfc.setResizable(false);

        Container c = sfc.getContentPane();
        c.setLayout(new FlowLayout());

        // Labels
        JPanel panel1 = new JPanel();
        JLabel directoryLabel = new JLabel("Choose search directory: ");

        final JTextField chosenDirectory = new JTextField(39);
        JButton dirButton = new JButton("...");

        JPanel panel2 = new JPanel();
        JLabel outputLabel = new JLabel("Choose output location:   ");
        final JTextField chosenOutput = new JTextField(39);
        JButton outButton = new JButton("...");

        JPanel panel3 = new JPanel();
        JCheckBox recursive = new JCheckBox("Search subdirectories?", true);
        JCheckBox hyperlink = new JCheckBox("Hyperlink to PDF path?", true);
        JCheckBox relative_link = new JCheckBox("Relative hyperlink?", true);

        JButton runButton = new JButton("Run");

        JButton openCSV = new JButton("Open CSV");

        JButton openDir = new JButton("Open search directory");

        // Layout
        JPanel panel4 = new JPanel();
        journal.setColumns(65);
        journal.setRows(20);
        journal.update(journal.getGraphics());

        JPanel panel5 = new JPanel();
        progress.setValue(0);
        progress.setMinimum(0);
        progress.setMaximum(100);
        progress.setStringPainted(true);
        progress.setString("");

        panel1.add(directoryLabel);
        panel1.add(chosenDirectory);
        panel1.add(dirButton);

        panel2.add(outputLabel);
        panel2.add(chosenOutput);
        panel2.add(outButton);

        panel3.add(recursive);
        panel3.add(hyperlink);
        panel3.add(relative_link);
        panel3.add(runButton);

        panel4.add(sp);

        panel5.add(progress);
        panel5.add(openCSV);
        panel5.add(openDir);

        c.add(panel1);
        c.add(panel2);
        c.add(panel3);
        c.add(panel4);
        c.add(panel5);


        // Styling

        Color background = new Color(70,80,70);
        Color foreground = new Color(255,230,250);
        Font labelfont = new Font("Courier", Font.PLAIN, 18);
        Font contentfont = new Font("Courier", Font.PLAIN, 14);
        c.setBackground(background);
        try { 
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            e.printStackTrace();
        }

        journal.setLineWrap(true);
        journal.setWrapStyleWord(true);

        // build GUI
        JPanel[] panels = {panel1, panel2, panel3, panel4, panel5};
        for (JPanel jPanel : panels) {
            jPanel.setBackground(background);
        }

        JLabel[] labels = {directoryLabel, outputLabel};
        for (JLabel jLabel : labels) {
            jLabel.setFont(labelfont);
            jLabel.setForeground(foreground);
        }

        JButton[] jButtons = {dirButton, outButton, runButton, openCSV, openDir};
        for (JButton jButton : jButtons) {
            jButton.setFont(contentfont);
        }

        JCheckBox[] jCheckBoxs = {recursive, hyperlink, relative_link};
        for (JCheckBox jCheckBox : jCheckBoxs) {
            jCheckBox.setFont(contentfont);
        }

        // Listeners 
        recursive.addItemListener(e -> {
            File dir = new File(chosenDirectory.getText());
            PDF_List = listPDF(Objects.requireNonNull(dir.listFiles()), recursive.isSelected());
            journal.setText("Found " + PDF_List.length + " PDFs\nPress run to extract metadata"); // refresh pdf number
                                                                                                  // for recursive
                                                                                                  // option
        });

        dirButton.addActionListener(ae -> {
            JFileChooser chooser = new JFileChooser();
            chooser.setMultiSelectionEnabled(true);
            chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
            if (chosenDirectory.getText() != null) {
                chooser.setCurrentDirectory(new File(chosenDirectory.getText()));
            }
            Action details = chooser.getActionMap().get("viewTypeDetails");
            details.actionPerformed(null);

            int option = chooser.showOpenDialog(sfc);
            if (option == JFileChooser.APPROVE_OPTION) {

                chosenDirectory.setText(String.valueOf(chooser.getSelectedFile()));

                File dir = new File(chosenDirectory.getText());

                PDF_List = listPDF(Objects.requireNonNull(dir.listFiles()), recursive.isSelected());

                journal.setText("Found " + PDF_List.length + " PDFs\nPress run to extract metadata");

                chosenOutput.setText(chooser.getSelectedFile() + "\\PDF_Metadata.csv");
            }

        });

        outButton.addActionListener(ae -> {
            JFileChooser chooser = new JFileChooser();
            chooser.setMultiSelectionEnabled(false);
            chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
            Action details = chooser.getActionMap().get("viewTypeDetails");
            if (chosenDirectory.getText() != null) {
                chooser.setCurrentDirectory(new File(chosenDirectory.getText()));
            }
            details.actionPerformed(null);

            int option = chooser.showOpenDialog(sfc);
            if (option == JFileChooser.APPROVE_OPTION) {
                chosenOutput.setText(chooser.getSelectedFile() + "\\PDF_Metadata.csv"); // could add default file name
                                                                                        // option and change string to
                                                                                        // variable
            }

        });

        openCSV.addActionListener(ae -> {

            try {
                Desktop.getDesktop().open(new File(chosenOutput.getText()));
            } catch (IOException e) {
                printJournal("Error while opening CSV");
                throw new RuntimeException(e);
            }
        });

        openDir.addActionListener(ae -> {

            try {
                Desktop.getDesktop().open(new File(chosenDirectory.getText()));
            } catch (IOException e) {
                printJournal("error while opening output directory");
                throw new RuntimeException(e);
            }
        });

        runButton.addActionListener(ae -> {

            if (PDF_List.length > 0) {
                String[][] Metadata_Array;
                try {
                    Metadata_Array = generate(PDF_List, hyperlink.isSelected(), chosenDirectory.getText());
                    if (hyperlink.isSelected() && relative_link.isSelected()) {
                        Path first = Paths.get(chosenDirectory.getText());
                        Path second = Paths.get(chosenOutput.getText());
                        String relativePath = String.valueOf(second.relativize(first));
                        for (int i = 0; i < Metadata_Array.length; i++) {
                            Metadata_Array[i][0] = Metadata_Array[i][0].replace((chosenDirectory.getText() + "\\"),
                                    relativePath.substring(0, relativePath.length() - 2)); // substring removes the
                                                                                           // final /.. in relativePath
                                                                                           // which escapes from the
                                                                                           // file name in chosenOutput
                        }
                    }
                } catch (IOException e) {
                    printJournal("Unknown error while generating csv");
                    throw new RuntimeException(e);
                }

                BufferedWriter br;
                StringBuilder sb = new StringBuilder();

                // Append strings from array
                for (String[] row : Metadata_Array) {
                    for (String element : row) {
                        if (element != null && !element.equals("null")) {
                            String element_string = "\"" + element + "\"";
                            sb.append(element_string);
                        }
                        sb.append(",");
                    }
                    sb.append("\n");
                }

                try {
                    br = new BufferedWriter(new FileWriter(chosenOutput.getText()));
                    br.write(sb.toString());
                    br.close();
                } catch (IOException e) {
                    printJournal("Error: output destination is in use by another program");
                    throw new RuntimeException(e);
                }

                progress.setString("\nFinished");
            } else {
                journal.append("\nNo PDFs to analyze");
            }

        });

        sfc.setVisible(true);
    }

    public static void printJournal(String msg) {
        journal.append("\n" + msg);
        journal.update(journal.getGraphics());
    }

    public static void progress(int current, int total) {
        float pct = (float) current / total;
        pct = pct * 100;
        progress.setValue((int) pct);
        progress.setString((int) pct + "%");
        progress.update(progress.getGraphics());
    }

    public static String[] listPDF(File[] files, boolean recursive) {
        // I didn't know about DFS when i wrote this
        String[] pdf_list = {};

        String[] directory_list = {};

        for (File file : files) {
            if (file.isDirectory()) { // list subdirectories in folder
                if (recursive) {

                    String[] New_Dir = new String[directory_list.length + 1];
                    int i;
                    for (i = 0; i < directory_list.length; i++) {
                        New_Dir[i] = directory_list[i];
                    }

                    New_Dir[i] = file.getPath();

                    directory_list = New_Dir;

                }

            } else {

                if (file.getPath().substring(file.getPath().length() - 3).equalsIgnoreCase("pdf")) { // this limits the metadata 
                                                                                                     // gathering to PDFs 
                                                                                                     // Could expand to collect at least 
                                                                                                     // doc info from all filetypes

                    String[] New_PDF = new String[pdf_list.length + 1];
                    int i;
                    for (i = 0; i < pdf_list.length; i++) {
                        New_PDF[i] = pdf_list[i];
                    }

                    New_PDF[i] = file.getAbsolutePath();

                    pdf_list = New_PDF;

                }
            }
        }
        int i;
        for (i = 0; i < directory_list.length; i++) { // for all subdirectories repeat process and append the pdfs in
                                                      // that subdirectory to the previous layer's pdf list
            File file = new File(directory_list[i]);

            String[] pdfs_in_subdirectory = listPDF(Objects.requireNonNull(file.listFiles()), true); // if this is
                                                                                                     // called recursive
                                                                                                     // must be true

            String[] New_PDF = new String[pdf_list.length + pdfs_in_subdirectory.length];
            int j;
            for (j = 0; j < pdf_list.length; j++) {
                New_PDF[j] = pdf_list[j];
            }
            for (j = 0; j < pdfs_in_subdirectory.length; j++) {
                New_PDF[j + pdf_list.length] = pdfs_in_subdirectory[j];
            }
            pdf_list = New_PDF;
        }

        return pdf_list;
    }

    public static String[][] generate(String[] PDF_List, boolean hyperlink, String Dir) throws IOException {

        String[][] Metadata_Array = { { "File Path",
                "Title", "Author", "Creator", "Creation Date", "Modification Date", "Producer", "Subject", "Keywords",
                "Trapped",
                "Title", "Description", "Creators", "Dates", "Contributors", "Coverage", "Description Languages",
                "Format", "Identifier", "Languages", "Publishers", "Relationships", "Rights", "Rights Languages",
                "Source", "Subjects", "Title Languages", "Types", "About", "Element",
                "Keywords", "PDF Version", "PDF Producer", "About", "Element",
                "Title", "Create Date", "Modify Date", "Metadata Date", "Creator Tool", "Advisories", "Base URL",
                "Identifiers", "Label", "Nickname", "Rating", "Thumbnail", "Thumbnail Languages", "About", "Element"
        } };

        Boolean output_requested = false;
        if (new File(Dir + "/pdf_metadata_output/").exists()) {
            output_requested = true;
        }

        for (String PDF : PDF_List) {
            PDDocument document;

            try {
                document = PDDocument.load(new File(PDF));
            } catch (IOException e) {
                printJournal("Open file error on " + PDF);
                continue;
            }

            // PDF = PDF.replaceAll("[^\\x00-\\x7F]"," "); //this replaces special
            // characters to stop excel opening the csv with improper encoding. Not
            // necessary if you use Excel 'import from csv' which is recommended.

            String location = PDF;
            if (hyperlink) {
                location = "=hyperlink(\"\"" + PDF + "\"\")";
            }

            PDDocumentCatalog catalog = document.getDocumentCatalog();

            PDMetadata meta = catalog.getMetadata();

            XMPSchemaDublinCore dc = null;
            XMPSchemaPDF pdf = null;
            XMPSchemaBasic basic = null;

            PDDocumentInformation information = document.getDocumentInformation();

            if (meta != null) {
                try {
                    XMPMetadata metadata = XMPMetadata.load(meta.exportXMPMetadata());

                    dc = metadata.getDublinCoreSchema();
                    pdf = metadata.getPDFSchema();
                    basic = metadata.getBasicSchema();
                } catch (IOException e) {
                    printJournal("XML parse error on " + PDF);
                    document.close();
                }
            }

            String[][] PDF_Metadata = new String[Metadata_Array.length + 1][Metadata_Array[0].length];

            int i;
            int j;
            for (i = 0; i < Metadata_Array.length; i++) {

                for (j = 0; j < Metadata_Array[0].length; j++) {

                    PDF_Metadata[i][j] = Metadata_Array[i][j];

                }
            }

            if (i % 5 == 0 || i == PDF_List.length) {
                progress(i, PDF_List.length);
            }

            // location
            PDF_Metadata[i][0] = display(location);
            // document information
            PDF_Metadata[i][1] = display(information.getTitle());
            PDF_Metadata[i][2] = display(information.getAuthor());
            PDF_Metadata[i][3] = display(information.getCreator());
            PDF_Metadata[i][4] = display(information.getCreationDate());

            PDF_Metadata[i][5] = display(information.getModificationDate());
            PDF_Metadata[i][6] = display(information.getProducer());
            PDF_Metadata[i][7] = display(information.getSubject());
            PDF_Metadata[i][8] = display(information.getKeywords());
            PDF_Metadata[i][9] = display(information.getTrapped());
            // Dublin Core
            if (dc != null) {
                PDF_Metadata[i][10] = display(dc.getTitle());
                PDF_Metadata[i][11] = display(dc.getDescription());
                PDF_Metadata[i][12] = listToString(dc.getCreators());
                PDF_Metadata[i][13] = listToString(dc.getDates());
                PDF_Metadata[i][14] = listToString(dc.getContributors());
                PDF_Metadata[i][15] = display(dc.getCoverage());
                PDF_Metadata[i][16] = listToString(dc.getDescriptionLanguages());
                PDF_Metadata[i][17] = display(dc.getFormat());
                PDF_Metadata[i][18] = display(dc.getIdentifier());
                PDF_Metadata[i][19] = listToString(dc.getLanguages());
                PDF_Metadata[i][20] = listToString(dc.getPublishers());
                PDF_Metadata[i][21] = listToString(dc.getRelationships());
                PDF_Metadata[i][22] = display(dc.getRights());
                PDF_Metadata[i][23] = listToString(dc.getRightsLanguages());
                PDF_Metadata[i][24] = display(dc.getSource());
                PDF_Metadata[i][25] = listToString(dc.getSubjects());
                PDF_Metadata[i][26] = listToString(dc.getTitleLanguages());
                PDF_Metadata[i][27] = listToString(dc.getTypes());
                PDF_Metadata[i][28] = display(dc.getAbout());
                PDF_Metadata[i][29] = display(dc.getElement());
            }
            // PDF
            if (pdf != null) {
                PDF_Metadata[i][30] = display(pdf.getKeywords());
                PDF_Metadata[i][31] = display(pdf.getPDFVersion());
                PDF_Metadata[i][32] = display(pdf.getProducer());
                PDF_Metadata[i][33] = display(pdf.getAbout());
                PDF_Metadata[i][34] = display(pdf.getElement());
            }
            // Basic
            if (basic != null) {
                PDF_Metadata[i][35] = display(basic.getTitle());
                PDF_Metadata[i][36] = display(basic.getCreateDate());
                PDF_Metadata[i][37] = display(basic.getModifyDate());
                PDF_Metadata[i][38] = display(basic.getMetadataDate());
                PDF_Metadata[i][39] = display(basic.getCreatorTool());
                PDF_Metadata[i][40] = listToString(basic.getAdvisories());
                PDF_Metadata[i][41] = display(basic.getBaseURL());
                PDF_Metadata[i][42] = listToString(basic.getIdentifiers());
                PDF_Metadata[i][43] = display(basic.getLabel());
                PDF_Metadata[i][44] = display(basic.getNickname());
                PDF_Metadata[i][45] = display(basic.getRating());
                PDF_Metadata[i][46] = display(basic.getThumbnail());
                PDF_Metadata[i][47] = listToString(basic.getThumbnailLanguages());
                PDF_Metadata[i][48] = display(basic.getAbout());
                PDF_Metadata[i][49] = display(basic.getElement());
            }

            Metadata_Array = PDF_Metadata;

            // this sections adds a page summarizing the pdf doc info to the start of each
            // pdf
            // this conveniently labels the metadata on the first page of each document
            // it is useful for printing and physically reviewing metadata patterns.
            // will only run if the folder "pdf_metadata_output" exists in the search
            // directory

            if (output_requested) {
                PDPage Metadata_Page = new PDPage();

                PDFont font = PDType1Font.HELVETICA_BOLD;

                // write the metadata to the page
                PDPageContentStream contentStream = new PDPageContentStream(document, Metadata_Page);

                for (int item = 0; item < 7; item++) {
                    contentStream.beginText();
                    contentStream.setFont(font, 12);
                    contentStream.newLineAtOffset(25, 700 - (25 * item));
                    if (item == 0) {
                        contentStream.showText(Metadata_Array[0][item].toString() + ": " + PDF);

                        // hyperlink to original file -- this will always be the absolute path
                        PDRectangle position = new PDRectangle();
                        PDAnnotationLink txtLink = new PDAnnotationLink();
                        PDActionURI action = new PDActionURI();
                        action.setURI("file://" + PDF.replace("\\", "/"));
                        txtLink.setAction(action);
                        position.setLowerLeftX(20);
                        position.setLowerLeftY(690);
                        position.setUpperRightX(590);
                        position.setUpperRightY(720);
                        txtLink.setAction(action);
                        txtLink.setRectangle(position);
                        Metadata_Page.getAnnotations().add(txtLink);

                    } else {
                        contentStream.showText(Metadata_Array[0][item].toString() + ":  " + PDF_Metadata[i][item]);
                    }
                    contentStream.endText();
                }

                contentStream.close();

                // add page to start of pdf
                PDPageTree pages = document.getDocumentCatalog().getPages();
                pages.insertBefore(Metadata_Page, pages.get(0));

                document.save(Dir + "/pdf_metadata_output/" + "output (" + i + ").pdf");
            }

            document.close();
        }

        return Metadata_Array;

    }

    private static String format(Object o) {

        if (o instanceof Calendar) {
            Calendar cal = (Calendar) o;
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSSXXX");
            sdf.setTimeZone(TimeZone.getTimeZone("GMT-05:00"));
            return sdf.format(cal.getTime());
        } else {
            return o.toString();
        }
    }

    private static String display(Object value) // String title,
    {
        if (value != null) {
            return format(value);
        } else {
            return "null";
        }
    }

    private static String listToString(List<?> list) {
        if (list == null) {
            return "null"; // not technically null but imported into Excel as null with the import tool
        }

        StringBuilder listString = new StringBuilder();

        int i = 1;

        for (Object item : list) {
            if (item == null) {
                continue;
            }
            item = format(item);
            if (i == 1) {
                listString.append(item);
            } else {
                listString.append(",").append(item);
            }
            i++;
        }

        return listString.toString();
    }

}
