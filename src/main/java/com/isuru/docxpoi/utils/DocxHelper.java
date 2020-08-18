package com.isuru.docxpoi.utils;

import com.isuru.docxpoi.dto.UserDto;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.stereotype.Component;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
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
        HashMap<String, String> map = new HashMap<String, String>();
        map.put(VariableTypes.FIRST_NAME.toString(), dto.getFirstName());
        map.put(VariableTypes.LAST_NAME.toString(), dto.getLastName());
        doc = replaceTextFor(doc, map);
        saveWord("src/main/resources/document.docx", doc);
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

    private void saveWord(String filePath, XWPFDocument doc) throws FileNotFoundException, IOException {
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
