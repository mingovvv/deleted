package itsm.excelmergetest;

import lombok.RequiredArgsConstructor;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.nio.charset.StandardCharsets;

@RequiredArgsConstructor
@RestController
@RequestMapping("/export-excel")
public class ExcelController {

    private static final MediaType APPLICATION_XLSX = MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

    private final ExcelService ExcelService;

    @GetMapping("/test")
    public ResponseEntity<ByteArrayResource> downloadExcel() {
        return ResponseEntity.ok()
                .contentType(APPLICATION_XLSX)
                .header(HttpHeaders.CONTENT_DISPOSITION, ContentDisposition.attachment().filename("test.xlsx", StandardCharsets.UTF_8).build().toString())
                .body(new ByteArrayResource(ExcelService.downloadExcel()));

    }

}
