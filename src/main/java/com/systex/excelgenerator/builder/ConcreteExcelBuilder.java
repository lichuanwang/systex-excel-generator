package com.systex.excelgenerator.builder;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.model.Candidate;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;

public class ConcreteExcelBuilder extends ExcelBuilder {

    private Candidate candidate;

    public ConcreteExcelBuilder(Candidate candidate) {
        this.candidate = candidate;
    }

    @Override
    public void buildHeader() {
        // Build header logic if necessary
    }

    @Override
    public void buildSections() throws IOException {
        XSSFSheet sheet = excelFile.createSheet("Candidate Information");

        // 測試內部連結可以導到同個excel file的不同sheet
        XSSFSheet test_sheet = excelFile.createSheet("Test");

        Section personalInfoSection = new PersonalInfoSection(candidate);
        Section educationSection = new EducationSection(candidate.getEducationList());
        Section experienceSection = new ExperienceSection(candidate.getExperienceList());
        Section projectSection = new ProjectSection(candidate.getProjects());
        Section SkillSection = new SkillSection(candidate.getSkills());

        // 只要new一個物件..? 因為可以存取不同的candidate..?

        int rowNum = 0;
        rowNum = personalInfoSection.populate(sheet, rowNum);
        rowNum += 5;

        rowNum = educationSection.populate(sheet, rowNum);
        rowNum += candidate.getEducationList().size() + 3;

        rowNum = experienceSection.populate(sheet, rowNum);
        rowNum += candidate.getExperienceList().size() + 3;

        rowNum = projectSection.populate(sheet, rowNum);
        rowNum += candidate.getProjects().size() + 3;

        rowNum= SkillSection.populate(sheet, rowNum);
        System.out.println(rowNum);

    }

    @Override
    public void buildFooter() {
        // Build footer logic if necessary
    }
}
