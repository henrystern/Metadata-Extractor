import org.apache.jempbox.xmp.XMPMetadata;
import org.apache.jempbox.xmp.XMPSchemaBasic;
import org.apache.jempbox.xmp.XMPSchemaDublinCore;
import org.apache.jempbox.xmp.XMPSchemaPDF;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentCatalog;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;
import org.apache.pdfbox.pdmodel.common.PDMetadata;

import javax.swing.*;
import java.awt.*;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.util.Calendar;
import java.util.List;
import java.util.Objects;

import java.nio.file.Path;
import java.nio.file.Paths;

public class MetadataExtractor {

    private static final JTextArea journal = new JTextArea(); //where messages are output
    private static final JProgressBar progress = new JProgressBar();

    private static String[] PDF_List = new String[0];

    public static void main(String[] args) {

        System.out.println("Running");

        JFrame sfc = new JFrame("PDF Metadata Extractor");

        sfc.setSize(800, 600);
        sfc.setPreferredSize(new Dimension(800,600));
        sfc.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

        Container c = sfc.getContentPane();
        c.setLayout(new FlowLayout());

        JPanel panel1 = new JPanel();
        JLabel directoryLabel = new JLabel("Choose search directory: ");
        final JTextField chosenDirectory = new JTextField(39);
        JButton dirButton = new JButton("...");


        JPanel panel2 = new JPanel();
        JLabel outputLabel = new JLabel("Choose output location: ");
        final JTextField chosenOutput = new JTextField(39);
        JButton outButton = new JButton("...");

        JPanel panel3 = new JPanel();
        JCheckBox recursive = new JCheckBox("Search subdirectories?", true);
        JCheckBox hyperlink = new JCheckBox("Hyperlink to PDF path?", true);
        JCheckBox relative_link = new JCheckBox("Relative hyperlink?", true);

        JButton runButton = new JButton("Run");

        JButton openCSV = new JButton("Open CSV");

        JButton openDir = new JButton("Open search directory");


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

        panel4.add(journal);

        panel5.add(progress);
        panel5.add(openCSV);
        panel5.add(openDir);

        c.add(panel1);
        c.add(panel2);
        c.add(panel3);
        c.add(panel4);
        c.add(panel5);

        recursive.addItemListener(e -> {
            File dir = new File(chosenDirectory.getText());
            PDF_List = listPDF(Objects.requireNonNull(dir.listFiles()), recursive.isSelected());
            journal.setText("Found " + PDF_List.length + " PDFs\nPress run to extract metadata"); //refresh pdf number for recursive option
        });


        dirButton.addActionListener(ae -> {
            JFileChooser chooser = new JFileChooser();
            chooser.setMultiSelectionEnabled(true);
            chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
            if (chosenDirectory.getText() != null){
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
            if (chosenDirectory.getText() != null){
                chooser.setCurrentDirectory(new File(chosenDirectory.getText()));
            }
            details.actionPerformed(null);


            int option = chooser.showOpenDialog(sfc);
            if (option == JFileChooser.APPROVE_OPTION) {
                chosenOutput.setText(chooser.getSelectedFile() + "\\PDF_Metadata.csv"); //could add default file name option and change string to variable
            }

        });

        openCSV.addActionListener(ae -> {

            try {
                Desktop.getDesktop().open(new File(chosenOutput.getText()));
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        });

        openDir.addActionListener(ae -> {

            try {
                Desktop.getDesktop().open(new File(chosenDirectory.getText()));
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        });

        runButton.addActionListener(ae -> {

            if (PDF_List.length > 0) {
                String[][] Metadata_Array;
                try {
                    Metadata_Array = generate(PDF_List, hyperlink.isSelected());
                    if(hyperlink.isSelected() && relative_link.isSelected()){
                        Path first = Paths.get(chosenDirectory.getText()); Path second = Paths.get(chosenOutput.getText());
                        String relativePath = String.valueOf(second.relativize(first));
                        for(int i=0; i < Metadata_Array.length; i++){
                            Metadata_Array[i][0] = Metadata_Array[i][0].replace((chosenDirectory.getText() + "\\"), relativePath.substring(0, relativePath.length() - 2)); //substring removes the final /.. in relativePath which escapes from the file name in chosenOutput
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

    public static void printJournal(String msg){
        journal.append("\n"+msg);
        journal.update(journal.getGraphics());
    }

    public static void progress(int current, int total){
        float pct = (float) current / total;
        pct = pct * 100;
        progress.setValue((int) pct);
        progress.setString((int) pct + "%");
        progress.update(progress.getGraphics());
    }

    public static String[] listPDF(File[] files, boolean recursive) {
        String[] pdf_list = {};

        String[] directory_list = {};

        for (File file : files) {
            if (file.isDirectory()) { //list subdirectories in folder
                if(recursive){

                    String[] New_Dir = new String[directory_list.length + 1];
                    int i;
                    for(i = 0; i < directory_list.length; i++) {
                        New_Dir[i] = directory_list[i];
                    }

                    New_Dir[i] = file.getPath();

                    directory_list = New_Dir;

                }

            } else {

                if (file.getPath().substring(file.getPath().length() - 3).equalsIgnoreCase("pdf")) {

                    String[] New_PDF = new String[pdf_list.length + 1];
                    int i;
                    for(i = 0; i < pdf_list.length; i++) {
                        New_PDF[i] = pdf_list[i];
                    }

                    New_PDF[i] = file.getAbsolutePath();

                    pdf_list = New_PDF;

                }
            }
        }

        int i;
        for(i = 0; i < directory_list.length; i++){ //for all subdirectories repeat process and append the pdfs in that subdirectory to the previous layer's pdf list
            File file = new File(directory_list[i]);

            String[] pdfs_in_subdirectory = listPDF(Objects.requireNonNull(file.listFiles()), true); //if this is called recursive must be true

            String[] New_PDF = new String[pdf_list.length + pdfs_in_subdirectory.length];
            int j;
            for(j = 0; j < pdf_list.length; j++) {
                New_PDF[j] = pdf_list[j];
            }
            for(j = 0; j < pdfs_in_subdirectory.length; j++) {
                New_PDF[j + pdf_list.length] = pdfs_in_subdirectory[j];
            }
            pdf_list = New_PDF;
        }

        return pdf_list;
    }

    public static String[][] generate(String[] PDF_List, boolean hyperlink) throws IOException {

        String[][] Metadata_Array = { {"File Path",
                "Title","Author","Creator","Creation Date","Modification Date","Producer","Subject","Keywords","Trapped",
                "Title","Description","Creators","Dates","Contributors","Coverage","Description Languages","Format","Identifier","Languages","Publishers","Relationships","Rights","Rights Languages","Source","Subjects","Title Languages","Types","About","Element",
                "Keywords","PDF Version","PDF Producer","About","Element",
                "Title","Create Date","Modify Date","Metadata Date","Creator Tool","Advisories","Base URL","Identifiers","Label","Nickname","Rating","Thumbnail","Thumbnail Languages","About","Element"
        } };



        for (String PDF : PDF_List) {
            PDDocument document;

            try {
                document = PDDocument.load(new File(PDF));
            }
            catch (IOException e) {
                printJournal("Open file error on " + PDF);
                continue;
            }

            //PDF = PDF.replaceAll("[^\\x00-\\x7F]"," "); //this replaces special characters to stop excel opening the csv with improper encoding. Not necessary if you use Excel 'import from csv' which is recommended.

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
                }
                catch (IOException e) {
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

            //location
            PDF_Metadata[i][0] = display(location);
            //document information
            PDF_Metadata[i][1] = display(information.getTitle());
            PDF_Metadata[i][2] = display(information.getAuthor());
            PDF_Metadata[i][3] = display(information.getCreator());
            PDF_Metadata[i][4] = display(information.getCreationDate());
            PDF_Metadata[i][5] = display(information.getModificationDate());
            PDF_Metadata[i][6] = display(information.getProducer());
            PDF_Metadata[i][7] = display(information.getSubject());
            PDF_Metadata[i][8] = display(information.getKeywords());
            PDF_Metadata[i][9] = display(information.getTrapped());
            //Dublin Core
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
            //PDF
            if (pdf != null) {
                PDF_Metadata[i][30] = display(pdf.getKeywords());
                PDF_Metadata[i][31] = display(pdf.getPDFVersion());
                PDF_Metadata[i][32] = display(pdf.getProducer());
                PDF_Metadata[i][33] = display(pdf.getAbout());
                PDF_Metadata[i][34] = display(pdf.getElement());
            }
            //Basic
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

            document.close();
        }

        return Metadata_Array;

    }

    private static String format(Object o)
    {

        if (o instanceof Calendar)
        {
            Calendar cal = (Calendar)o;
            return DateFormat.getDateInstance().format(cal.getTime());
        }
        else
        {
            return o.toString();
        }
    }

    private static String display(Object value) //String title,
    {
        if (value != null)
        {
            return format(value);
        }
        else
        {
            return "null";
        }
    }

    private static String listToString(List<?> list)
    {
        if (list == null)
        {
            return "null"; //not technically null but easy enough to replace with ctrl+h in Excel
        }

        StringBuilder listString = new StringBuilder();

        for (Object item : list ) {
            if (item == null){
                continue;
            }
            item = format(item);
            listString.append(",").append(item);
        }

        return listString.toString();
    }

}
