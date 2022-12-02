package TestCases;

import Utilities.ExcelUtils;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class Iteration2 {
    public WebDriver driver;

    @Test
    public void main() throws InterruptedException, IOException {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("http://nstar-web-lazio-tsi.apps.osv1.aci.it/");

        driver.findElement(By.xpath("//a[contains(text(),'ACCEDI')]")).click();

        //login
        WebElement username = driver.findElement(By.xpath("//input[@id=\"username\"]"));
        username.sendKeys("g.miranda");

        WebElement password = driver.findElement(By.xpath("//input[@id=\"password\"]"));
        password.sendKeys("iniziale");

        Thread.sleep(3000);

        driver.findElement(By.xpath("//input[@id=\"kc-login\"]")).click();

        Thread.sleep(4000);


        //selzione Calcolo Tariffa
        driver.findElement(By.xpath("(//a[@class='dropdown-toggle nav-link'])[1]")).click();

        driver.findElement(By.xpath("(//a[contains(text(),' Calcolo Tariffa ')])[1]")).click();

        Thread.sleep(2000);


        XSSFWorkbook workbook;
        XSSFSheet sheet;
        XSSFCell cell;

        // Import excel sheet.
        File src = new File("src/test/resources/Tariffario_NSTAR_2022_Umbria2.xlsx");

        // Load the file.
        FileInputStream fis = new FileInputStream(src);

        // Load the workbook.
        workbook = new XSSFWorkbook(fis);
        //Load the sheet in which data is stored.
        sheet = workbook.getSheet("Tariffario_NSTAR_2022_Umbria");


        /** Reload Excel*/

        ExcelUtils file = new ExcelUtils("src/test/resources/Tariffario_NSTAR_2022_Umbria2.xlsx", "Tariffario_NSTAR_2022_Umbria");

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            List<Map<String, String>> dataList = file.getDataList();
            String[] data = new String[0];
            for (Map<String, String> oneRow : dataList) {
                String CLass = oneRow.get("Classe");
                String KW = oneRow.get("KW");
                String Uso = oneRow.get("Uso");
                String Alimentazione = oneRow.get("Alim");
                String Euro = oneRow.get("Euro");
                String Potenza = oneRow.get("Potenza");
                String Cilindrata = oneRow.get("Cilindrata");


                // if (!CLass.isBlank() && !KW.isBlank() && !Uso.isBlank() ) {
               // System.out.println(CLass + "___" + KW + Uso + Euro);

                /** Inserimento Categoria */

                Select objSelect = new Select(driver.findElement(By.xpath("//select[@class=\"custom-select ng-untouched ng-pristine ng-invalid\"]")));
                objSelect.selectByValue(CLass);
                Thread.sleep(1000);



                 /*I have added test data in the cell A2 as "testemailone@test.com" and B2 as "password"
               Cell A2 = row 1 and column 0. It reads first row as 0, second row as 1 and so on
               and first column (A) as 0 and second column (B) as 1 and so on*/

                /** Data Valadità */
                cell = sheet.getRow(1).getCell(3);
                cell.setCellType(CellType.STRING);
                Thread.sleep(2000);
                driver.findElement(By.xpath("//input[@id=\"dataValidita\"]")).sendKeys(cell.getStringCellValue());


                /** Data Immatricolazione */
                cell = sheet.getRow(1).getCell(9);
                cell.setCellType(CellType.STRING);
                Thread.sleep(2000);
                driver.findElement(By.xpath("//input[@id=\"dataImmatricolazione\"]")).sendKeys(cell.getStringCellValue());


                /** Mesi */
                WebElement twelve2 = driver.findElement(By.xpath("//input[@id=\"numeroMesi\"]"));
                twelve2.click();
                twelve2.sendKeys("12");


                /**Uso */
                Select uso = new Select(driver.findElement(By.xpath("//select[@id=\"uso\"]")));
                uso.selectByValue(Uso);
                Thread.sleep(1000);


                /** Alimentazione */
                Select alimentazione = new Select(driver.findElement(By.xpath("//select[@id=\"alimentazione\"]")));
                alimentazione.selectByValue(Alimentazione);
                Thread.sleep(1000);

                WebElement dropDown1 = driver.findElement(By.id("elementId"));
                Select selectElement = new Select(dropDown1);
                selectElement.selectByVisibleText("OptionText");
                selectElement.deselectAll();

                Select DropDown = new Select(driver.findElement(By.id("Drp_ID")));

                DropDown.deselectAll();
            }

             continue;




        }
    }
}