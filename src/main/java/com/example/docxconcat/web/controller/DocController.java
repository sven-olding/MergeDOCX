package com.example.docxconcat.web.controller;

import com.example.docxconcat.service.DocxMerger;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.Objects;

/**
 * @author ogbozoyan
 * @date 08.07.2023
 */
@RestController
@CrossOrigin(origins = "*", allowedHeaders = "*")
@Slf4j
@RequestMapping("/doc")
@Tag(name = "Docx controller", description = "Controller to merge docx")
public class DocController {

    @Autowired
    private DocxMerger docxMerger;

    @Operation(summary = "Returns a merged file")
    @PostMapping(
            value = "/concat",
            consumes = MediaType.MULTIPART_FORM_DATA_VALUE,
            produces = MediaType.APPLICATION_OCTET_STREAM_VALUE
    )
    public ResponseEntity<ByteArrayResource> concat(@RequestParam("sourceDoc") MultipartFile sourceDoc, @RequestParam("docToAdd") MultipartFile docToAdd) throws Exception {
        MultipartFile mergedDoc;
        try {
            mergedDoc = docxMerger.merge(sourceDoc.getInputStream(), docToAdd.getInputStream());

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData(Objects.requireNonNull(mergedDoc.getOriginalFilename()), mergedDoc.getOriginalFilename());
            headers.setContentLength(mergedDoc.getSize());

            ByteArrayResource resource;

            resource = new ByteArrayResource(mergedDoc.getBytes());
            log.debug("File successfully merged");
            return new ResponseEntity<>(resource, headers, HttpStatus.OK);
        } catch (IOException e) {
            log.debug("File could not be merged ", e);
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}
