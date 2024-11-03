package com.systex.excelgenerator.builder;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Candidate;
import com.systex.excelgenerator.model.Project;
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
    public void buildBody() {
        ExcelSheet sheet = excelFile.createSheet("Candidate Information");

        PersonalInfoSection personalInfoSection = new PersonalInfoSection();
        personalInfoSection.setData(candidate);

        EducationSection educationSection = new EducationSection();
        educationSection.setData(candidate.getEducationList());

        ExperienceSection experienceSection = new ExperienceSection();
        experienceSection.setData(candidate.getExperienceList());

        ProjectSection projectSection = new ProjectSection();
        projectSection.setData(candidate.getProjects());

        SkillSection skillSection = new SkillSection();
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
