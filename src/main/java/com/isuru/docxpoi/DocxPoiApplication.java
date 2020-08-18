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
import java.util.HashMap;

@SpringBootApplication
public class DocxPoiApplication {

    public static void main(String[] args) {
        SpringApplication.run(DocxPoiApplication.class, args);
    }


}



