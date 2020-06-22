package app.controller;

import app.service.ExcelAuditService;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.PostConstruct;
import java.io.File;
import java.io.IOException;
import java.util.Objects;

@Controller
public class ExcelAuditController {

    private final ExcelAuditService auditService;

    public ExcelAuditController(ExcelAuditService auditService) {
        this.auditService = auditService;
    }

    @PostMapping("/upload")
    public ResponseEntity<String> uploadFile(@RequestParam MultipartFile file) {

        String[] filename = Objects.requireNonNull(file.getOriginalFilename()).toLowerCase().split("\\.");
        if (filename[filename.length-1].startsWith("xl")) {
            System.out.println("Wait until application process your file");
            try {
                auditService.processFile(file.getBytes());
            } catch (IOException | InvalidFormatException e) {
                e.printStackTrace();
                return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
            }
        } else {
            return ResponseEntity.status(HttpStatus.NOT_ACCEPTABLE).body("File extension " + filename[filename.length-1] + " is not supported");
        }
        return ResponseEntity.ok().build();
    }

}
