package com.isuru.docxpoi;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

@SpringBootApplication
public class DocxPoiApplication {

    public static void main(String[] args) throws URISyntaxException, IOException {
        SpringApplication.run(DocxPoiApplication.class, args);

        String resourcePath = "template.docx";
        Path templatePath = Paths.get(DocxPoiApplication.class.getClassLoader().getResource(resourcePath).toURI());
        XWPFDocument doc = new XWPFDocument(Files.newInputStream(templatePath));
        doc = replaceTextFor(doc, "UNIQUE_VAR", "MyValue1");
        saveWord("document.docx", doc);
    }

    private static XWPFDocument replaceTextFor(XWPFDocument doc, String findText, String replaceText) {
        doc.getParagraphs().forEach(p -> {
            p.getRuns().forEach(run -> {
                String text = run.text();
                if (text.contains(findText)) {
                    run.setText(text.replace(findText, replaceText), 0);
                }
            });
        });

        return doc;
    }

    private static void saveWord(String filePath, XWPFDocument doc) throws FileNotFoundException, IOException {
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(filePath);
            doc.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            out.close();
        }
    }

}



