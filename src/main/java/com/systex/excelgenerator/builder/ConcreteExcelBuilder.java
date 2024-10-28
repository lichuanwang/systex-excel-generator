package com.systex.excelgenerator.builder;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.model.Candidate;
import org.apache.poi.xssf.usermodel.XSSFSheet;

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
    public void buildSections() {
        XSSFSheet sheet = excelFile.createSheet("Candidate Information");

        Section personalInfoSection = new PersonalInfoSection(candidate);
        Section educationSection = new EducationSection(candidate.getEducationList());
        Section experienceSection = new ExperienceSection(candidate.getExperienceList());
        Section projectSection = new ProjectSection(candidate.getProjects());
        Section SkillSection = new SkillSection(candidate.getSkills());

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
