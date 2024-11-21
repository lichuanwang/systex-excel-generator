package com.systex.excelgenerator;

import com.systex.excelgenerator.model.*;
import com.systex.excelgenerator.service.ExcelGenerationService;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

public class Main {
    private static final Logger log = LogManager.getLogger(Main.class);

    public static void main(String[] args) {
        // Step 1: Create a mock Candidate with sample data
        Candidate candidate1 = new Candidate();
        candidate1.setName("John Doe");
        candidate1.setEmail("john.doe@gmail.com");
        candidate1.setPhone("0123456789");
        candidate1.setBirthday(new Date(1999));
        candidate1.setGender("Male");

        Address address1 = new Address();
        address1.setStreet("123 Main St");
        address1.setCity("Springfield");
        address1.setZip("62704");
        candidate1.setAddress(address1);

        // Add Education data for Candidate 1
        Education education1 = new Education();
        education1.setSchoolName("Springfield University");
        education1.setMajor("Bachelor of Science in Computer Science");
        education1.setStartDate(LocalDate.of(2019, 9, 30));
        education1.setEndDate(LocalDate.of(2024, 6, 30));
        candidate1.setEducationList(Arrays.asList(education1));

        // Add Experience data for Candidate 1
        Experience experience1 = new Experience();
        experience1.setCompanyName("Tech Solutions Inc.");
        experience1.setJobTitle("Software Engineer");
        experience1.setStartDate(LocalDate.of(2019, 9, 30));
        experience1.setEndDate(LocalDate.of(2020, 10, 30));
        experience1.setDescription("Developed large scale application");
        candidate1.setExperienceList(Arrays.asList(experience1));

        // Add Skills data for Candidate 1
        Skill skill1 = new Skill();
        skill1.setId(1);
        skill1.setSkillName("Java");
        skill1.setLevel(5);
        Skill skill2 = new Skill();
        skill2.setId(2);
        skill2.setSkillName("Spring Boot");
        skill2.setLevel(2);
        Skill skill3 = new Skill();
        skill3.setId(3);
        skill3.setSkillName("Angular");
        skill3.setLevel(3);

        candidate1.setSkills(Arrays.asList(skill1, skill2, skill3));

        // Step 5: Add Projects data
        Project project1 = new Project();
        project1.setProjectName("E-commerce Platform");
        project1.setDescription("Developed an online shopping platform with Spring Boot and React.Developed an online shopping platform with Spring Boot and React.Developed an online shopping platform with Spring Boot and React. with Spring Boot and React.Developed an online shopping pltaaaaaaBoot anssssss.");
        project1.setRole("Web Developer");
        project1.setTechnologiesUsed("Angular, Spring Boot");
        candidate1.setProjects(Arrays.asList(project1));

        // Add ImagePath for Candidate 1
        List<String> imagepathList = new ArrayList<>();
        imagepathList.add("C:\\Users\\2400823\\Downloads\\test.jpg");
        imagepathList.add("C:\\Users\\2400823\\Downloads\\SSU_Kirby_artwork.png");
        candidate1.setImagepath(imagepathList);

        // Step 2: Create another mock Candidate with different data
        Candidate candidate2 = new Candidate();
        candidate2.setName("Jane Smith");
        candidate2.setEmail("jane.smith@gmail.com");
        candidate2.setPhone("0987654321");
        candidate2.setBirthday(new Date(1997));
        candidate2.setGender("Female");

        Address address2 = new Address();
        address2.setStreet("456 Elm St");
        address2.setCity("Metropolis");
        address2.setZip("12345");
        candidate2.setAddress(address2);

        // Add Education data for Candidate 2
        Education education2 = new Education();
        education2.setSchoolName("Metropolis University");
        education2.setMajor("Master of Business Administration");
        education2.setStartDate(LocalDate.of(2021, 9, 30));
        education2.setEndDate(LocalDate.of(2023, 6, 30));
        candidate2.setEducationList(Arrays.asList(education2));

        // Add Experience data for Candidate 2
        Experience experience2 = new Experience();
        experience2.setCompanyName("Business Corp.");
        experience2.setJobTitle("Business Analyst");
        experience2.setStartDate(LocalDate.of(2021, 1, 15));
        experience2.setEndDate(LocalDate.of(2023, 5, 30));
        experience2.setDescription("Analyzed business requirements and developed solutions");
        candidate2.setExperienceList(Arrays.asList(experience2));

        // Add Skills data for Candidate 2
        Skill skill4 = new Skill();
        skill4.setId(2);
        skill4.setSkillName("Data Analysis");
        Skill skill5 = new Skill();
        skill5.setId(3);
        skill5.setSkillName("Project Management");
        candidate2.setSkills(Arrays.asList(skill2, skill3));

        // Add Projects data for Candidate 2
        Project project2 = new Project();
        project2.setProjectName("Market Analysis Platform");
        project2.setDescription("Led the development of a platform for market data analysis.Led the development of a platform for market data analysis.");
        project2.setRole("Team Lead");
        project2.setTechnologiesUsed("Python, Tableau");
        candidate2.setProjects(Arrays.asList(project2));

        // Step 3: Add both candidates to the candidate list
        List<Candidate> candidateList = new ArrayList<>();
        candidateList.add(candidate1);
        candidateList.add(candidate2);

        // Step 4: Generate Excel file for both candidates
        ExcelGenerationService excelGenerationService = new ExcelGenerationService();
        excelGenerationService.generateExcelForCandidate(candidateList);

        log.info("Excel file generated successfully!");
    }
}


