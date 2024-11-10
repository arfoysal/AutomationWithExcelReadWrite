package org.example;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;
import pages.BasePage;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

public class AppTest {

    public WebDriver driver;
    @BeforeMethod
    public WebDriver getDriver(){
        driver = new ChromeDriver();
        driver.get("https://www.google.com/");
        return driver;
    }

    @AfterMethod
    public void closeDriver(){
        driver.quit();
    }
    BasePage basePage = new BasePage();
    @Test
    public void testDay() throws InterruptedException {
        String dayName = basePage.getDayName();

        Sheet sheet = basePage.getSheet(dayName);
        List<String> searchData = basePage.getDaySearchData(sheet);


        for (String data: searchData){
            driver.findElement(By.name("q")).clear();
            driver.findElement(By.name("q")).sendKeys(data);
            Thread.sleep(3000);

            List<WebElement> list_result = driver.findElements(By.xpath("//ul/li[@role='presentation']//div[@role='option']/div[1]"));
            System.out.println("Searching for " + data);
            List<String> result =basePage.getShortestAndLongestText(list_result);
            basePage.writeResultData(dayName, data, result);

        }
    }
}
