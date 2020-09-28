package com.isuru.docxpoi.utils;

import com.isuru.docxpoi.dto.UserDto;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.stereotype.Component;

import java.io.*;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;

@Component
public class DocxHelper {

    public void createDocx(UserDto dto) throws URISyntaxException, IOException {

        String resourcePath = "template.docx";
        Path templatePath = Paths.get(DocxHelper.class.getClassLoader().getResource(resourcePath).toURI());
        XWPFDocument doc = new XWPFDocument(Files.newInputStream(templatePath));
        HashMap<String, String> map = new HashMap<>();
        map.put(VariableTypes.FIRST_NAME.toString(), dto.getFirstName());
        map.put(VariableTypes.LAST_NAME.toString(), dto.getLastName());
        doc = replaceTextFor(doc, map);
        savePdf("src/main/resources/document.pdf", doc);
    }

    private XWPFDocument replaceTextFor(XWPFDocument doc, HashMap map) {
        doc.getParagraphs().forEach(p -> {
            p.getRuns().forEach(run -> {
                String text = run.text();
                map.forEach((findText, replaceText) -> {
                    if (text.contains((String) findText)) {
                        run.setText(text.replace((String) findText, (String) replaceText), 0);
                    }
                });
            });
        });

        return doc;
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
}
