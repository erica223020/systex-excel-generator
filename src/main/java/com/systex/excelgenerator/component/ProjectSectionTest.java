package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Project;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class ProjectSectionTest extends AbstractSection<Project> {

    private List<Project> projects;
    private CellStyle headerStyle;
    private CellStyle dataStyle;

    public ProjectSectionTest(XSSFWorkbook workbook) {
        super("Project",workbook);
        initializeStyles();
    }

    private void initializeStyles() {
        headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerStyle.setFont(headerFont);

        dataStyle = workbook.createCellStyle();
        Font dataFont = workbook.createFont();
        dataFont.setFontHeightInPoints((short) 12);
        dataStyle.setFont(dataFont);
    }

    public CellStyle getHeaderStyle() {
        return headerStyle;
    }

    public CellStyle getDataStyle() {
        return dataStyle;
    }

    @Override
    protected int generateHeader(XSSFSheet sheet, int rowNum) {
        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("Project");
        headerRow.createCell(0).setCellStyle(headerStyle);
        headerRow.createCell(1).setCellValue("Role");
        headerRow.createCell(1).setCellStyle(headerStyle);
        headerRow.createCell(2).setCellValue("Description");
        headerRow.createCell(2).setCellStyle(headerStyle);
        headerRow.createCell(3).setCellValue("Technology");
        headerRow.createCell(3).setCellStyle(headerStyle);
        return rowNum;
    }

    @Override
    protected int generateData(XSSFSheet sheet, int rowNum) {
        for (Project project : projects) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(project.getProjectName());
            row.getCell(0).setCellStyle(dataStyle);
            row.createCell(1).setCellValue(project.getRole());
            row.getCell(1).setCellStyle(dataStyle);
            row.createCell(2).setCellValue(project.getDescription());
            row.getCell(2).setCellStyle(dataStyle);
            row.createCell(3).setCellValue(project.getTechnologiesUsed());
            row.getCell(3).setCellStyle(dataStyle);
        }
        return rowNum;
    }

    @Override
    protected int generateFooter(XSSFSheet sheet, int rowNum) {
        return rowNum;
    }

    @Override
    public void setData(Project data) {
        this.projects = Arrays.asList(data);
    }

    @Override
    public void setData(Collection<Project> dataCollection) {
        if (dataCollection != null && !dataCollection.isEmpty()) {
            this.projects = new ArrayList<>(dataCollection);
        }
    }

    @Override
    public boolean isEmpty() {
         return projects == null || projects.isEmpty();
    }
}
