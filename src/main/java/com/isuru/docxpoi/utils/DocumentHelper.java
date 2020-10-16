package com.isuru.docxpoi.utils;

import com.isuru.docxpoi.dto.EmployeeDetails;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.springframework.stereotype.Component;

import java.io.*;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;

@Component
public class DocumentHelper {

    public ByteArrayOutputStream createDocument(Integer id) throws URISyntaxException, IOException {

        //get employee by id from database. this is dummy data.
        EmployeeDetails employeeDetails = EmployeeDetails.builder()
                .firstName("Ranil")
                .lastName("Perera")
                .address("100,Temple Street,Colombo")
                .dob(LocalDate.now().minusYears(26))
                .employeeId(id)
                .gender("Male")
                .position("Software Engineer")
                .build();

        String resourcePath = "template.docx";
        Path templatePath = Paths.get(DocumentHelper.class.getClassLoader().getResource(resourcePath).toURI());
        XWPFDocument doc = new XWPFDocument(Files.newInputStream(templatePath));

        HashMap<String, String> map = new HashMap<>();
        map.put(VariableTypes.FIRST_NAME.getName(), employeeDetails.getFirstName());
        map.put(VariableTypes.LAST_NAME.getName(), employeeDetails.getLastName());
        map.put(VariableTypes.POSITION.getName(), employeeDetails.getPosition());
        map.put(VariableTypes.GENDER.getName(), employeeDetails.getGender());
        map.put(VariableTypes.DATE_OF_BIRTH.getName(), employeeDetails.getDob().toString());
        map.put(VariableTypes.ADDRESS.getName(), employeeDetails.getAddress());
        map.put(VariableTypes.EMPLOYEE_ID.getName(), employeeDetails.getEmployeeId().toString());

        replaceTextFor(doc, map);

        //get data from database. this is dummy data.
        List<SalaryRecord> salaryRecordList = Arrays.asList(
                SalaryRecord.builder().month("Jan 2020").amount(String.valueOf(1200.30)).build(),
                SalaryRecord.builder().month("Feb 2020").amount(String.valueOf(1200.30)).build(),
                SalaryRecord.builder().month("Mar 2020").amount(String.valueOf(1200.30)).build(),
                SalaryRecord.builder().month("Apr 2020").amount(String.valueOf(1200.30)).build(),
                SalaryRecord.builder().month("May 2020").amount(String.valueOf(1500.70)).build(),
                SalaryRecord.builder().month("Jun 2020").amount(String.valueOf(1500.70)).build()
        );

        replaceSalaryTable(doc, salaryRecordList);

        savePdf("src/main/resources/employee-report.pdf", doc);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        PdfOptions options = PdfOptions.create();
        PdfConverter.getInstance().convert(doc, out, options);
        out.close();
        return out;
    }

    private void replaceTextFor(XWPFDocument doc, HashMap map) {
        doc.getParagraphs().forEach(p -> p.getRuns().forEach(run -> {
            String text = run.text();
            map.forEach((findText, replaceText) -> {
                if (text.contains((String) findText)) {
                    run.setText(text.replace((String) findText, (String) replaceText), 0);
                }
            });
        }));
    }

    private void savePdf(String filePath, XWPFDocument doc) {
        try {
            PdfOptions options = PdfOptions.create();
            OutputStream out = new FileOutputStream(new File(filePath));
            PdfConverter.getInstance().convert(doc, out, options);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void replaceSalaryTable(XWPFDocument doc, List<SalaryRecord> salaryRecordList) {
        XWPFTable table = doc.getTableArray(0);
        int templateRowId = 1;
        XWPFTableRow rowTemplate = table.getRow(templateRowId);

        salaryRecordList.forEach(salaryRecord -> {

            CTRow ctrow = null;
            try {
                ctrow = CTRow.Factory.parse(rowTemplate.getCtRow().newInputStream());
            } catch (XmlException | IOException e) {
                e.printStackTrace();
            }

            XWPFTableRow newRow = new XWPFTableRow(ctrow, table);

            newRow.getCell(0).getParagraphArray(0).getRuns().get(0).setText(salaryRecord.getMonth(), 0);
            newRow.getCell(1).getParagraphArray(0).getRuns().get(0).setText(salaryRecord.getAmount(), 0);

            table.addRow(newRow);
        });

        table.removeRow(templateRowId);
    }
}