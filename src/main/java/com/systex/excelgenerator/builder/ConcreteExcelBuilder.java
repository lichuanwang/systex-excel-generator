package com.systex.excelgenerator.builder;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Candidate;

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
    public void buildBody() {
        ExcelSheet sheet = excelFile.createSheet("Candidate Information");

        PersonalInfoDataSection personalInfoSection = new PersonalInfoDataSection();
        personalInfoSection.setData(candidate);

        EducationDataSection educationSection = new EducationDataSection();
        educationSection.setData(candidate.getEducationList());

        ExperienceDataSection experienceSection = new ExperienceDataSection();
        experienceSection.setData(candidate.getExperienceList());

        ProjectDataSection projectSection = new ProjectDataSection();
        projectSection.setData(candidate.getProjects());

        SkillDataSection skillSection = new SkillDataSection();
        skillSection.setData(candidate.getSkills());


        if (!personalInfoSection.isEmpty()) {
            sheet.addSection(personalInfoSection);
        }

        if (!educationSection.isEmpty()) {
            sheet.addSection(educationSection);
        }

        if (!experienceSection.isEmpty()) {
            sheet.addSection(experienceSection);
        }

        if (!projectSection.isEmpty()) {
            sheet.addSection(projectSection);
        }

        if (!skillSection.isEmpty()) {
            sheet.addSection(skillSection);
        }
    }

//    @Override
//    public void buildBody() {
//        ExcelSheet sheet = excelFile.createSheet("Candidate Information");
//
//        // Create new section
//        EducationSection educationSection = new EducationSection();
////        PersonalInfoSection personalInfoSection = new PersonalInfoSection();
////        ExperienceSection experienceSection = new ExperienceSection();
////        ProjectSection projectSection = new ProjectSection();
////        SkillSection skillSection = new SkillSection();
//
//        // Assign data to each section
//        educationSection.setData(candidate.getEducationList());
////        personalInfoSection.setData(candidate);
////        experienceSection.setData(candidate.getExperienceList());
////        projectSection.setData(candidate.getProjects());
////        skillSection.setData(candidate.getSkills());
//
//
////        if (!personalInfoSection.isEmpty()) {
////            personalInfoSection.render(sheet);
////        }
//
//        if (!educationSection.isEmpty()) {
//            educationSection.render(sheet);
//        }
////
////        if (!experienceSection.isEmpty()) {
////            rowNum = experienceSection.populate(sheet, rowNum);
////            rowNum += 5;
////        }
////
////        if (!projectSection.isEmpty()) {
////            rowNum = projectSection.populate(sheet, rowNum);
////            rowNum += 5;
////        }
////
////        if (!skillSection.isEmpty()) {
////            rowNum = skillSection.populate(sheet, rowNum);
////            rowNum += 5;
////        }
////
////        System.out.println(rowNum);
//    }

    @Override
    public void buildFooter() {
        // Build footer logic if necessary
    }
}
