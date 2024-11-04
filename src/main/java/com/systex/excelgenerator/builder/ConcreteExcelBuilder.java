package com.systex.excelgenerator.builder;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.model.Candidate;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ConcreteExcelBuilder extends ExcelBuilder {

    private Candidate candidate;
    private XSSFWorkbook workbook; // 將 workbook 定義為 XSSFWorkbook 類型

    public ConcreteExcelBuilder(Candidate candidate) {
        this.candidate = candidate;
        this.workbook = new XSSFWorkbook(); // 初始化 workbook
    }

    @Override
    public void buildHeader() {
        // Build header logic if necessary
    }

    @Override
    public void buildSections() {
        // 使用 workbook 創建 XSSFSheet
        XSSFSheet sheet = workbook.createSheet("Candidate Information");

        // 使用 workbook 傳入每個 Section
        PersonalInfoSection personalInfoSection = new PersonalInfoSection(workbook);
        EducationSection educationSection = new EducationSection(workbook);
        ExperienceSection experienceSection = new ExperienceSection(workbook);
        ProjectSection projectSection = new ProjectSection(workbook);
        SkillSection skillSection = new SkillSection(workbook);

        // Assign data to each section
        personalInfoSection.setData(candidate);
        educationSection.setData(candidate.getEducationList());
        experienceSection.setData(candidate.getExperienceList());
        projectSection.setData(candidate.getProjects());
        skillSection.setData(candidate.getSkills());

        int rowNum = 0;
        if (!personalInfoSection.isEmpty()) {
            rowNum = personalInfoSection.populate(sheet, rowNum);
            rowNum += 5;
        }

        if (!educationSection.isEmpty()) {
            rowNum = educationSection.populate(sheet, rowNum);
            rowNum += 5;
        }

        if (!experienceSection.isEmpty()) {
            rowNum = experienceSection.populate(sheet, rowNum);
            rowNum += 5;
        }

        if (!projectSection.isEmpty()) {
            rowNum = projectSection.populate(sheet, rowNum);
            rowNum += 5;
        }

        if (!skillSection.isEmpty()) {
            rowNum = skillSection.populate(sheet, rowNum);
            rowNum += 5;
        }

        System.out.println(rowNum);
    }

    @Override
    public void buildFooter() {
        // Build footer logic if necessary
    }

    // Getter for workbook to use later if needed
    public XSSFWorkbook getWorkbook() {
        return workbook;
    }
    public void saveToFile(String filePath) {
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut); // 寫入 Excel 文件
            workbook.close(); // 關閉 workbook 資源
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

