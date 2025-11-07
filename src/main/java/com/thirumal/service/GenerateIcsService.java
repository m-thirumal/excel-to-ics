package com.thirumal.service;

import static net.fortuna.ical4j.model.property.immutable.ImmutableCalScale.GREGORIAN;
import static net.fortuna.ical4j.model.property.immutable.ImmutableVersion.VERSION_2_0;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.thirumal.Holiday;

import net.fortuna.ical4j.data.CalendarOutputter;
import net.fortuna.ical4j.model.Calendar;
import net.fortuna.ical4j.model.ComponentList;
import net.fortuna.ical4j.model.PropertyList;
import net.fortuna.ical4j.model.component.VEvent;
import net.fortuna.ical4j.model.property.Categories;
import net.fortuna.ical4j.model.property.Description;
import net.fortuna.ical4j.model.property.DtStart;
import net.fortuna.ical4j.model.property.ProdId;
import net.fortuna.ical4j.model.property.Status;
import net.fortuna.ical4j.model.property.Summary;
import net.fortuna.ical4j.model.property.Transp;
import net.fortuna.ical4j.model.property.Uid;
import net.fortuna.ical4j.model.property.XProperty;

/**
 * @author ThirumalM
 */
@Service
public class GenerateIcsService {
	
	Logger logger = LoggerFactory.getLogger(GenerateIcsService.class);

	public File generate(MultipartFile file) throws Exception {
		Map<String, List<Holiday>> branchHolidays = readExcel(file);

        // If only one branch → return one .ics file
        if (branchHolidays.size() == 1) {
            Map.Entry<String, List<Holiday>> entry = branchHolidays.entrySet().iterator().next();
            return createIcsFile(entry.getKey(), entry.getValue());
        }

        // Otherwise create ZIP
        File zipFile = new File("Company_Holidays_" + System.currentTimeMillis() + ".zip");
        try (FileOutputStream fos = new FileOutputStream(zipFile);
             ZipOutputStream zipOut = new ZipOutputStream(fos)) {

            for (Map.Entry<String, List<Holiday>> entry : branchHolidays.entrySet()) {
                File icsFile = createIcsFile(entry.getKey(), entry.getValue());
                try (FileInputStream fis = new FileInputStream(icsFile)) {
                    ZipEntry zipEntry = new ZipEntry(icsFile.getName());
                    zipOut.putNextEntry(zipEntry);
                    fis.transferTo(zipOut);
                }
            }
        }
        return zipFile;
	}
	
	private Map<String, List<Holiday>> readExcel(MultipartFile file) throws IOException {
        Map<String, List<Holiday>> map = new LinkedHashMap<>();

        try (InputStream is = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(is)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    logger.warn("Skipping empty row at index {}", i);
                    continue;
                }
                boolean isRowEmpty = true;
                for (int j = 0; j <= 3; j++) {
                    if (row.getCell(j) != null && !row.getCell(j).toString().trim().isEmpty()) {
                        isRowEmpty = false;
                        break;
                    }
                }
                if (isRowEmpty) {
                    logger.warn("Skipping completely empty row at index {}", i);
                    continue;
                }
                logger.debug("Row : {}", row);
                LocalDate date = row.getCell(0).getLocalDateTimeCellValue().toLocalDate();
                String branch = row.getCell(1).getStringCellValue().trim();
                String summary = row.getCell(2).getStringCellValue().trim();
                String desc = (row.getCell(3) != null) ? row.getCell(3).getStringCellValue().trim() : "";

                map.computeIfAbsent(branch, k -> new ArrayList<>())
                        .add(new Holiday(date, branch, summary, desc));
            }
        }
        return map;
    }

    private File createIcsFile(String branch, List<Holiday> holidays) throws Exception {
        PropertyList properties = new PropertyList();
        properties.add(new ProdId("-//NeSL//" + branch + " Holidays 2026//EN"));
        properties.add(VERSION_2_0);
        properties.add(GREGORIAN);
        properties.add(new XProperty("X-WR-CALNAME", branch + " Holidays 2026"));
        properties.add(new XProperty("X-WR-TIMEZONE", "Asia/Kolkata"));
        
        Calendar calendar = new Calendar(properties, new ComponentList());
        
        for (Holiday h : holidays) {
        	logger.info("Adding holiday: {} - {} ({})", h.getDate(), h.getSummary(), branch);
            // Define required properties
            DtStart<LocalDate> start = new DtStart<>(h.getDate());
         //   start.getParameters().add(Value.DATE); // ensures it's treated as all-day event
            //
            Summary summary = new Summary(h.getSummary());
            Description desc = new Description(h.getDescription());
            Uid uid = new Uid(UUID.randomUUID().toString());
            Categories categories = new Categories(branch);
            Status status = new Status(Status.VALUE_CONFIRMED);
            Transp transp = new Transp(Transp.VALUE_TRANSPARENT);

            // Create the event
            VEvent event = new VEvent(new PropertyList(), new ComponentList());
            event.getProperties().add(start);
            event.getProperties().add(summary);
            event.getProperties().add(desc);
            event.getProperties().add(uid);
            event.getProperties().add(categories);
            event.getProperties().add(status);
            event.getProperties().add(transp);

            calendar.getComponents().add(event);
        }


        File file = new File(branch.replaceAll("\\s+", "_") + "_Holidays.ics");
        try (FileOutputStream fout = new FileOutputStream(file)) {
            new CalendarOutputter().output(calendar, fout);
        }
        return file;
    }

}
