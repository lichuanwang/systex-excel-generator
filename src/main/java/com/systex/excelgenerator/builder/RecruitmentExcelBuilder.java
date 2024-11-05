package com.systex.excelgenerator.builder;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Candidate;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public class RecruitmentExcelBuilder extends ExcelBuilder {

    private Candidate candidate;

    public RecruitmentExcelBuilder(Candidate candidate) {
        this.candidate = candidate;
    }

    @Override
    public void buildHeader() {
        // Build header logic if necessary
    }

    @Override
    public void buildBody() {
        ExcelSheet sheet = excelFile.createSheet(candidate.getName());
        List<Section> sectionList = new ArrayList<>();

        initializeSection(sectionList, new PersonalInfoSection(), candidate);
        initializeSection(sectionList, new EducationSection(), candidate.getEducationList());
        initializeSection(sectionList, new ExperienceSection(), candidate.getExperienceList());
        initializeSection(sectionList, new ProjectSection(), candidate.getProjects());
        initializeSection(sectionList, new SkillSection(), candidate.getSkills());

        for (Section section : sectionList) {
            if (!section.isEmpty()) {
                sheet.addSection(section);
            }
        }
    }

    private void initializeSection(List<Section> sectionList, Section section, Object data) {
        section.setData(data);
        sectionList.add(section);
    }

    private <T> void initializeSection(List<Section> sectionList, Section section, Collection<T> data) {
        section.setData(data);
        sectionList.add(section);
    }

    @Override
    public void buildFooter() {
        // Build footer logic if necessary
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

