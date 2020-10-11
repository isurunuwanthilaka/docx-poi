package com.isuru.docxpoi.controller;

import com.isuru.docxpoi.dto.UserDto;
import com.isuru.docxpoi.utils.DocxHelper;
import org.apache.tomcat.util.file.ConfigurationSource;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayOutputStream;

@RestController
@RequestMapping("/v1")
public class DocxController {

    @Autowired
    private DocxHelper docxHelper;

    @GetMapping("/docx")
    private ResponseEntity createDocument() {
        try {
            UserDto dto = new UserDto();
            dto.setFirstName("Isuru");
            dto.setLastName("Nuwanthilaka");
            ByteArrayOutputStream out = docxHelper.createDocx(dto);
            ByteArrayResource resource = new ByteArrayResource(out.toByteArray());

            return ResponseEntity.ok()
                    .contentType(MediaType.parseMediaType("application/octet-stream"))
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + "testing.pdf" + "\"")
                    .body(out.toByteArray());
        } catch (Exception e) {
            return ResponseEntity.status(500).build();
        }
    }

}
