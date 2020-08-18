package com.isuru.docxpoi.controller;

import com.isuru.docxpoi.dto.UserDto;
import com.isuru.docxpoi.utils.DocxHelper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/v1")
public class DocxController {

    @Autowired
    private DocxHelper docxHelper;

    @PostMapping("/docx")
    private ResponseEntity createDocument(@RequestBody UserDto dto) {
        try {
            docxHelper.createDocx(dto);
            return ResponseEntity.ok().build();
        } catch (Exception e) {
            return ResponseEntity.status(500).build();
        }
    }

}
