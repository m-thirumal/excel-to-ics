/**
 * 
 */
package com.thirumal.controller;

import java.io.File;
import java.io.FileInputStream;

import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.thirumal.service.GenerateIcsService;

/**
 * @author ThirumalM
 */
@RequestMapping("/generate-ics")
@RestController
public class GenerateIcsController {

	private GenerateIcsService generateIcsService;
	
	public GenerateIcsController(GenerateIcsService generateIcsService) {
		super();
		this.generateIcsService = generateIcsService;
	}

	@PostMapping(consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
	public Object generate(@RequestParam MultipartFile file) {
		try {
            File generatedFile = generateIcsService.generate(file);
            InputStreamResource resource = new InputStreamResource(new FileInputStream(generatedFile));
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + generatedFile.getName() + "\"")
                    .contentType(MediaType.parseMediaType("text/calendar"))
                    .contentLength(generatedFile.length())
                    .body(resource);
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body(null);
        }
	}
	
}
