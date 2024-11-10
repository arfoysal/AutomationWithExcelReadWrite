package pages;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import java.io.*;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public class BasePage {

    public String getDayName(){
        // Get the current date
        LocalDate currentDate = LocalDate.now();

        // Get the day of the week
        String dayName = currentDate.getDayOfWeek().getDisplayName(TextStyle.FULL, Locale.ENGLISH);
        return dayName;
    }

    public Sheet getSheet(String dayName){
        File f = new File("src/test/resources");
        File data = new File(f, "DayKeywordData.xlsx");

        try(FileInputStream fis = new FileInputStream(data.getAbsolutePath());
            Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet(dayName);
            return sheet;

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public List<String> getDaySearchData(Sheet sheet){
        List<String> columnCValues = new ArrayList<>();
        for(Row row: sheet){
            Cell cell = row.getCell(2);
            if (cell.getCellType() == CellType.STRING){
                columnCValues.add(cell.getStringCellValue());
            }
        }
        return columnCValues;

    }

    public List<String> getShortestAndLongestText(List<WebElement> elements){
        List<String> data = new ArrayList<>();
        String defText = elements.get(0).getText();
        String shortest = defText;
        String longest = defText;
        for (WebElement element: elements){
            String text = element.getText();
            if (text.length() < shortest.length()){
                shortest = text;
            }
            if (text.length()> longest.length()){
                longest = text;
            }
        }
        data.add(longest);
        data.add(shortest);
        return data;
    }

    public void writeResultData(String dayName, String Keyword, List<String> values){
        File f = new File("src/test/resources");
        File data = new File(f, "DayKeywordData.xlsx");

        try(FileInputStream fis = new FileInputStream(data.getAbsolutePath());
            Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(dayName);
            for(Row row: sheet){
                Cell cell = row.getCell(2);
                if (cell.getCellType() == CellType.STRING){
                    if (cell.getStringCellValue().equals( Keyword)){
                        Cell longestCell = row.getCell(cell.getColumnIndex() + 1);
                        Cell shortestCell = row.getCell(cell.getColumnIndex() + 2);
                        longestCell.setCellValue(values.get(0));
                        shortestCell.setCellValue(values.get(1));
                        break;
                    }
                }
            }
            try (FileOutputStream fos = new FileOutputStream(data.getAbsolutePath())) {
                workbook.write(fos);
            }
            System.out.println("Longest result is " + values.get(0));
            System.out.println("Shortest result is " + values.get(1));
            System.out.println("Successfully updated the data for " + Keyword);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
