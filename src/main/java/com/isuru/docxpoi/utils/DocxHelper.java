package com.isuru.docxpoi.utils;

import com.isuru.docxpoi.dto.UserDto;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;

@Component
public class DocxHelper {

    public void createDocx(UserDto dto) throws URISyntaxException, IOException {

        String resourcePath = "template.docx";
        Path templatePath = Paths.get(DocxHelper.class.getClassLoader().getResource(resourcePath).toURI());
        XWPFDocument doc = new XWPFDocument(Files.newInputStream(templatePath));
        HashMap<String, String> map = new HashMap<>();
        map.put(VariableTypes.FIRST_NAME.toString(), dto.getFirstName());
        map.put(VariableTypes.LAST_NAME.toString(), dto.getLastName());
        replaceTextFor(doc, map);
        replaceTable(doc);
        savePdf("src/main/resources/document.pdf", doc);
    }

    private void replaceTextFor(XWPFDocument doc, HashMap map) {
        doc.getParagraphs().forEach(p -> p.getRuns().forEach(run -> {
            String text = run.text();
            map.forEach((findText, replaceText) -> {
                if (text.contains((String) findText)) {
                    run.setText(text.replace((String) findText, (String) replaceText), 0);
                }
            });
        }));
    }

    private void savePdf(String filePath, XWPFDocument doc) {
        try {
            PdfOptions options = PdfOptions.create();
            OutputStream out = new FileOutputStream(new File(filePath));
            PdfConverter.getInstance().convert(doc, out, options);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void replaceTable(XWPFDocument doc) {
        XWPFTable table = doc.getTableArray(0);
        int templateRowId = 1;
        XWPFTableRow rowTemplate = table.getRow(templateRowId);

        List<Data> reportData = Arrays.asList(
                Data.builder().col1("isuru").col2("nuwanthilaka").build(),
                Data.builder().col1("kamala").col2("nuwanthilaka").build()
        );

        reportData.forEach(data -> {

            CTRow ctrow = null;
            try {
                ctrow = CTRow.Factory.parse(rowTemplate.getCtRow().newInputStream());
            } catch (XmlException | IOException e) {
                e.printStackTrace();
            }

            XWPFTableRow newRow = new XWPFTableRow(ctrow, table);

            System.out.println(newRow.getCell(0).getParagraphArray(0).getRuns().get(0));

            newRow.getCell(0).getParagraphArray(0).getRuns().get(0).setText(data.getCol1(),0);
            newRow.getCell(1).getParagraphArray(0).getRuns().get(0).setText(data.getCol2(),0);

            table.addRow(newRow);
        });

        table.removeRow(templateRowId);
    }
}