package com.systex.excelgenerator.builder;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Candidate;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public class RecruitmentExcelBuilder extends ExcelBuilder {

    private final Candidate candidate;

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
        List<Section<?>> sectionList = new ArrayList<>();

        initializeSection(sectionList, new PersonalInfoSection(), candidate);
        initializeSection(sectionList, new EducationSection(), candidate.getEducationList());
        initializeSection(sectionList, new ExperienceSection(), candidate.getExperienceList());
        initializeSection(sectionList, new ProjectSection(), candidate.getProjects());
        initializeSection(sectionList, new SkillSection(), candidate.getSkills());

        for (Section<?> section : sectionList) {
            if (!section.isEmpty()) {
                sheet.addSection(section);
            }
        }
    }

    @Override
    public void buildFooter() {
        // Build footer logic if necessary
    }

    private <T> void initializeSection(List<Section<?>> sectionList, Section<T> section, T data) {
        section.setData(data);
        sectionList.add(section);
    }

    private <T> void initializeSection(List<Section<?>> sectionList, Section<T> section, Collection<T> data) {
        section.setData(data);
        sectionList.add(section);
    }
}

