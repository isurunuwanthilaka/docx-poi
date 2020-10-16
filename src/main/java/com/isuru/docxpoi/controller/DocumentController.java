package com.isuru.docxpoi.controller;

import com.isuru.docxpoi.utils.DocumentHelper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayOutputStream;

@RestController
@RequestMapping("/v1")
public class DocumentController {

    private final DocumentHelper documentHelper;

    @Autowired
    public DocumentController(DocumentHelper documentHelper) {
        this.documentHelper = documentHelper;
    }

    @GetMapping("report/employee/{id}")
    public ResponseEntity createDocument(@PathVariable("id") Integer id) {
        try {
            ByteArrayOutputStream out = documentHelper.createDocument(id);

            return ResponseEntity.ok()
                    .contentType(MediaType.parseMediaType("application/octet-stream"))
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"employee-report.pdf\"")
                    .body(out.toByteArray());
        } catch (Exception e) {
            return ResponseEntity.badRequest().build();
        }
    }

}
