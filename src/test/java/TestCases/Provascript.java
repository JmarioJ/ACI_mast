package TestCases;

import Utilities.ExcelUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import java.util.List;
import java.util.Map;

public class Provascript {

    public static void main(String[] args) throws InterruptedException {
        WebDriver driver = new ChromeDriver();
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

        Thread.sleep(4000);

        ExcelUtils file = new ExcelUtils("src/test/resources/Tariffario_NSTAR_2022_Umbria2.xlsx", "Tariffario_NSTAR_2022_Umbria");
        //how many rows in sheet
        //System.out.println("file.rowCount() = " + file.rowCount());
        //how many columns in sheet
        //System.out.println("file.columnCount() = " + file.columnCount());
        //get column names
        //System.out.println("file.getColumnsNames() = " + file.getColumnsNames());
        //
        List<Map<String, String>> dataList = file.getDataList();
        String[] data = new String[0];
        for (Map<String, String> oneRow : dataList) {
            String CLass = oneRow.get("Classe");
            String KW = oneRow.get("KW");

            if (!CLass.isBlank() && !KW.isBlank()) {
                System.out.println(CLass + "___" + KW);
                /** Inserimento Categotia */
                Select objSelect = new Select(driver.findElement(By.xpath("//select[@class=\"custom-select ng-untouched ng-pristine ng-invalid\"]")));
                objSelect.selectByValue(CLass);
                Thread.sleep(1000);

            }
            break;

            //Inserimento dati per il calcolo

            /** Inserimento Categotia
             Select objSelect = new Select(driver.findElement(By.xpath("//select[@class=\"custom-select ng-untouched ng-pristine ng-invalid\"]")));
             objSelect.selectByVisibleText("AUTOBUS");


             /** Data Validità */
            //WebElement DataValidità = driver.findElement(By.xpath("//input[@id=\"dataValidita\"]"));
            //DataValidità.sendKeys("01/01/2022");

            /** Data Immatricolazione */
            //WebElement DataImmatricolazione = driver.findElement(By.xpath("//input[@id=\"dataImmatricolazione\"]"));
            //DataImmatricolazione.sendKeys("01/01/2001");


            /** Mesi
             WebElement Mesi = driver.findElement(By.xpath("(//input[@class=\"form-control ng-untouched ng-pristine ng-valid\"])[1]"));
             Mesi.sendKeys("12");*/


            //Select  =new Select(driver.findElement(By.xpath("//select[@id=\"uso\"]")));
            //objSelect.selectByVisibleText("tttt");

            //Select objSelect= new Select(driver.findElement(By.xpath("//select[@id=\"uso\"]")));
            //objSelect.selectByVisibleText("");


        }

    }
}


