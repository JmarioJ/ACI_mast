package TestCases;

import Utilities.ExcelUtils;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.formula.functions.Value;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;
import org.testng.util.Strings;

import java.io.*;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

public class third {

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
        File src = new File("src/test/resources/Prpvarisultati.xlsx");

        // Load the file.
        FileInputStream fis = new FileInputStream(src);

        // Load the workbook.
        workbook = new XSSFWorkbook(fis);
        //Load the sheet in which data is stored.
        sheet = workbook.getSheet("Foglio1");


        /** Reload Excel*/

        ExcelUtils file = new ExcelUtils("src/test/resources/Prpvarisultati.xlsx", "Foglio1");

        String confirmationMessage2="";

       // String testData[][] = new String[][];


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
                String Portata = oneRow.get("Portata");
                String Peso = oneRow.get("Peso");
                String AssiMotrici = oneRow.get("AssiMotrici");
                String SospensioniPneumatiche = oneRow.get("SospensioniPneumatiche");
                String GancioTraino = oneRow.get("GancioTraino");
                String PesoRimorchio = oneRow.get("PesoRimorchio");


                // if (!CLass.isBlank() && !KW.isBlank() && !Uso.isBlank() ) {
                System.out.println(CLass + "___" + KW + Uso + Euro);

                /** Inserimento Categoria */
                Select objSelect = new Select(driver.findElement(By.xpath("//select[@class=\"custom-select ng-untouched ng-pristine ng-invalid\"]")));
                objSelect.selectByValue(CLass);
                Thread.sleep(1000);


                /** Data Valadità */
                cell = sheet.getRow(i).getCell(3);
                cell.setCellType(CellType.STRING);
                Thread.sleep(2000);
                driver.findElement(By.xpath("//input[@id=\"dataValidita\"]")).sendKeys(cell.getStringCellValue());


                /** Data Immatricolazione */
                cell = sheet.getRow(i).getCell(9);
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


                /** Potenza */
                cell = sheet.getRow(i).getCell(15);
                cell.setCellType(CellType.STRING);
                WebElement potenza = driver.findElement(By.xpath("//input[@id=\"kw\"]"));
                potenza.click();
                potenza.sendKeys(Potenza);


                /** Euro */
                Select Euro2 = new Select(driver.findElement(By.xpath("//select[@id=\"euro\"]")));
                Euro2.selectByValue(String.valueOf(Euro));


                /** if Portata does not have value skip, if it has value insert it in the application*/


                if ((Strings.isNullOrEmpty(Portata))) {
                    System.out.println("there is no any value " + i);

                } else if ((!Strings.isNullOrEmpty(Portata))) {
                    cell = sheet.getRow(i).getCell(20);
                    Thread.sleep(2000);
                    cell.setCellType(CellType.STRING);
                    Thread.sleep(2000);
                    driver.findElement(By.xpath("//input[@id=\"portata\"]")).sendKeys(cell.getStringCellValue());
                }


                /** insert Peso*/

                if ((Strings.isNullOrEmpty(Peso))) {
                    System.out.println("there is no any value " + i);

                } else if ((!Strings.isNullOrEmpty(Peso))) {
                    cell = sheet.getRow(i).getCell(21);
                    cell.setCellType(CellType.STRING);
                    Thread.sleep(2000);
                    driver.findElement(By.xpath("//input[@id=\"pesoComplessivo\"]")).sendKeys(cell.getStringCellValue());
                }

                /** insert AssiMotrici*/

                if ((Strings.isNullOrEmpty(AssiMotrici))) {
                    System.out.println("there is no any value " + i);

                } else if ((!Strings.isNullOrEmpty(AssiMotrici))) {
                    cell = sheet.getRow(i).getCell(22);
                    cell.setCellType(CellType.STRING);
                    Thread.sleep(2000);
                    driver.findElement(By.xpath("//input[@id=\"assiMotore\"]")).sendKeys(cell.getStringCellValue());
                }


                if ((Strings.isNullOrEmpty(PesoRimorchio))) {
                    System.out.println("there is no any value " + i);

                } else if ((!Strings.isNullOrEmpty(PesoRimorchio))) {
                    cell = sheet.getRow(i).getCell(26);
                    cell.setCellType(CellType.STRING);
                    Thread.sleep(2000);
                    driver.findElement(By.xpath("//input[@id=\"pesoRimorchio\"]")).sendKeys(cell.getStringCellValue());
                }


                if (SospensioniPneumatiche.contains("NO")) {
                    System.out.println("it did not click on check box ");

                } else if (SospensioniPneumatiche.contains("SI")) {
                    cell = sheet.getRow(i).getCell(24);
                    Thread.sleep(2000);
                    cell.setCellType(CellType.STRING);
                    Thread.sleep(2000);
                    WebElement sospensione = driver.findElement(By.xpath("//input[@id=\"sospensionePneumatica\"]"));
                    sospensione.click();
                } else if ((Strings.isNullOrEmpty(SospensioniPneumatiche))) {
                    System.out.println("it did not click on check box ");
                }


                if ((Strings.isNullOrEmpty(GancioTraino))) {
                    System.out.println("there is no any value " + i);

                } else if ((!Strings.isNullOrEmpty(GancioTraino))) {
                    cell = sheet.getRow(i).getCell(25);
                    cell.setCellType(CellType.STRING);
                    Thread.sleep(2000);
                    WebElement rimorchiabilità = driver.findElement(By.xpath("//input[@id=\"gancioTraino\"]"));
                    rimorchiabilità.click();
                }


                /**Cerca*/
                WebElement Cerca = driver.findElement(By.xpath("//button[@class=\"btn btn-primary\"]"));
                Cerca.click();


                driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);


                /** Cancella Filtri*/
                Thread.sleep(100);
                WebElement cancellafiltri = driver.findElement(By.xpath("//button[@class=\"btn btn-link\"]"));
                cancellafiltri.click();



                WebElement confirmationMessage = driver.findElement(By.xpath("//tbody/tr/td[13]"));
                 confirmationMessage2 = confirmationMessage.getText();

                 confirmationMessage2= confirmationMessage2.replace("€","").trim();


                confirmationMessage2= confirmationMessage2.replace(".",",").trim();
                 System.out.println(confirmationMessage2);


            }





            //To write into Excel File
            FileOutputStream outputStream = new FileOutputStream("src/test/resources/ACI_2.xlsx");
            workbook.write(outputStream);

            DataFormatter format = new DataFormatter();

            for (i = 1; i < sheet.getLastRowNum(); i++) {  //this is the array that stores our tab in excel

                if (confirmationMessage2 == format.formatCellValue(sheet.getRow(i).getCell(27)).trim()) {

                    XSSFCell cell2 = sheet.getRow(i).createCell(2);
                    System.out.println( format.formatCellValue(sheet.getRow(i).getCell(27)));
                    cell2.setCellValue("PASS");
                    System.out.println("Test Pass");
                    //append the result in a new array at column N.3

                } else {
                    XSSFCell cell2 = sheet.getRow(i).createCell(2);
                    System.out.println( format.formatCellValue(sheet.getRow(i).getCell(27)));
                    cell2.setCellValue("Fail");
                    System.out.println("Test Fail");

                    // append the result in a new array at column n.3
                }


// Basically we need three array:
// the first one is where you store the ExcelSheet TariffarioACI
// the second array will store the value you extract with the get text
//the third array will have the 1st column = to 1st column of the excel sheet
// second column exqual to EXPECTED RESULT from TARIFARRIO ACI
//third column equal to ACTUAL RESULT from second array
// fth column equal to the array that contains the result
            }
        }
    }
}


//format.formatCellValue(sheet.getRow(i).getCell(27)).trim())