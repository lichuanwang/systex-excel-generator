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

        // Create new section
        PersonalInfoSection personalInfoSection = new PersonalInfoSection();
        EducationSection educationSection = new EducationSection();
        ExperienceSection experienceSection = new ExperienceSection();
        ProjectSection projectSection = new ProjectSection();
        SkillSection skillSection = new SkillSection();

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
}
