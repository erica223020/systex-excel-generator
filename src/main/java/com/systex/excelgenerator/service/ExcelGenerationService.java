package com.systex.excelgenerator.service;

import com.systex.excelgenerator.builder.ConcreteExcelBuilder;
import com.systex.excelgenerator.model.Candidate;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelGenerationService {

    public void generateExcelForCandidate(Candidate candidate) {
        // 建立 Excel 內容
        ConcreteExcelBuilder builder = new ConcreteExcelBuilder(candidate);
        builder.buildHeader();
        builder.buildSections();
        builder.buildFooter();

        // 取得 workbook
        XSSFWorkbook workbook = builder.getWorkbook();

        // 儲存到檔案
        try (FileOutputStream fileOut = new FileOutputStream("candidate_info_test.xlsx")) {
            workbook.write(fileOut); // 寫入 Excel 文件
            workbook.close(); // 關閉 workbook 資源
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
