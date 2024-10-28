package com.systex.excelgenerator;


import com.systex.excelgenerator.model.*;
import com.systex.excelgenerator.service.ExcelGenerationService;

import java.time.LocalDate;
import java.util.Arrays;
import java.util.Date;

public class Main {
    public static void main(String[] args) {
        // Step 1: Create a mock Candidate with sample data
        Candidate candidate = new Candidate();
        candidate.setName("John Doe");
        candidate.setEmail("john.doe@gmail.com");
        candidate.setPhone("1234567890");
        candidate.setBirthday(new Date(1999));
        candidate.setGender("Male");

        Address address = new Address();
        address.setStreet("123 Main St");
        address.setCity("Springfield");
        address.setZip("62704");
        candidate.setAddress(address);

        // Step 2: Add Education data
        Education education1 = new Education();
        education1.setSchoolName("Springfield University");
        education1.setMajor("Bachelor of Science in Computer Science");
        education1.setStartDate(LocalDate.of(2019, 9, 30));
        education1.setEndDate(LocalDate.of(2024, 6, 30));

        candidate.setEducationList(Arrays.asList(education1));

        // Step 3: Add Experience data
        Experience experience1 = new Experience();
        experience1.setCompanyName("Tech Solutions Inc.");
        experience1.setJobTitle("Software Engineer");
        experience1.setStartDate(LocalDate.of(2019, 9, 30));
        experience1.setEndDate(LocalDate.of(2020, 10, 30));

        candidate.setExperienceList(Arrays.asList(experience1));

        // Step 4: Add Skills data
        Skill skill1 = new Skill();
        skill1.setId(1);
        skill1.setSkillName("Java");
        Skill skill2 = new Skill();
        skill2.setId(2);
        skill2.setSkillName("Spring Boot");

        candidate.setSkills(Arrays.asList(skill1, skill2));

        // Step 5: Add Projects data
        Project project1 = new Project();
        project1.setProjectName("E-commerce Platform");
        project1.setDescription("Developed an online shopping platform with Spring Boot and React.");
        project1.setRole("Web Developer");
        project1.setTechnologiesUsed("Angular, Spring Boot");

        candidate.setProjects(Arrays.asList(project1));

        // Step 6: Generate Excel file
        ExcelGenerationService excelGenerationService = new ExcelGenerationService();
        excelGenerationService.generateExcelForCandidate(candidate);

        System.out.println("Excel file generated successfully!");
    }
}
