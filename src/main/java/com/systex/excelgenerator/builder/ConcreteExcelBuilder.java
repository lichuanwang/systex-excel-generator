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

        personalInfoSection.populate(sheet);
        educationSection.populate(sheet);
        experienceSection.populate(sheet);
        projectSection.populate(sheet);
        SkillSection.populate(sheet);
    }

    @Override
    public void buildFooter() {
        // Build footer logic if necessary
    }
}
