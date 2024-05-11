import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
//import org.openqa.selenium.support.ui.Duration;
import java.time.Duration;



import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.concurrent.TimeUnit;

public class DataGettingFromExcel {

    public static void main(String[] args) throws InterruptedException, IOException {

//        System.setProperty("webDriver.chrome.driver", "C:\\Users\\Nitin\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.xlsx");
        System.setProperty("webDriver.chrome.driver", " C:\\Users\\Nitin\\Downloads\\chromedriver-win64\\chromedriver-win64.exe");
        WebDriver webDriver = new ChromeDriver();
        Actions actions = new Actions(webDriver);
//        WebDriver driver = DriverFactory.getDriverFor("chrome");
//        driver.manage().timeouts().implicitlyWait(30,TimeUnit.SECONDS);
        webDriver.get("http://192.168.0.221:7070");
        webDriver.manage().window().maximize();
        Thread.sleep(1000);
        webDriver.findElement(By.xpath("//*[@id='username']")).sendKeys("super.admin");
        webDriver.findElement(By.xpath("//*[@id=\":r1:\"]")).sendKeys("123456");
        Thread.sleep(1000);
        webDriver.findElement(By.xpath("//div[contains(text(),'Unit')]")).click();

        webDriver.findElement(By.xpath("//*[@id=\"react-select-2-option-0\"]")).click();
        webDriver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div/div[2]/div[2]/form/div[7]/button")).click();

        System.out.println("login successfully!");

        Thread.sleep(2000);

        // For Visit
        webDriver.findElement(By.cssSelector("button[aria-label='open drawer'] svg")).click();
        Thread.sleep(2000);
        webDriver.findElement(By.xpath("//p[normalize-space()='Registration']")).click();
       Thread.sleep(2000);
        webDriver.findElement(By.xpath("//p[normalize-space()='Quick Registration']")).click();
        actions.sendKeys(Keys.ENTER).build().perform();
        webDriver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div/div[1]/div/div/ul/div/div/div/div/li[2]/div/div[2]/a/div/span/p")).click();

        FileInputStream fis = new FileInputStream("/home/nitin/Documents/Sheet1.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        int rowCount = sheet.getLastRowNum();
        int columnCount = sheet.getRow(1).getLastCellNum();
        System.out.println("rowCount :" + rowCount + "columnCount :" + columnCount);
        int pinCodeXpathValue = 14;
        for (int i = 1; i <= rowCount; i++) {
            XSSFRow celldata = sheet.getRow(i);

            String firstName = celldata.getCell(0).getStringCellValue();
            if( firstName == null){

                break;

            }
            String middleName = celldata.getCell(1).getStringCellValue();

            String lastName = celldata.getCell(2).getStringCellValue();

            int age = (int) celldata.getCell(3).getNumericCellValue();

            String emailId = celldata.getCell(4).getStringCellValue();

            int mobile = (int)celldata.getCell(5).getNumericCellValue();
            int contactNo = (int) celldata.getCell(6).getNumericCellValue();

            String panNumber = celldata.getCell(7).getStringCellValue();

            String houseNo = celldata.getCell(8).getStringCellValue();

            String streetAddress = celldata.getCell(9).getStringCellValue();

            int pinCode = (int)celldata.getCell(10).getNumericCellValue();


//            int referenceNumber = (int)celldata.getCell(11).getNumericCellValue();
            //For Prefix
            webDriver.findElement(By.xpath("//div[contains(text(),'Prefix*')]")).click();
            actions.sendKeys(Keys.ARROW_DOWN).build().perform();
            actions.sendKeys(Keys.ENTER).build().perform();

            webDriver.findElement(By.xpath("//input[@name='firstName']")).clear();
            webDriver.findElement(By.xpath("//input[@name='firstName']")).sendKeys(firstName);
            webDriver.findElement(By.xpath("//input[@name='middleName']")).clear();
            webDriver.findElement(By.xpath("//input[@name='middleName']")).sendKeys(middleName);
            webDriver.findElement(By.xpath("//input[@name='lastName']")).clear();
            webDriver.findElement(By.xpath("//input[@name='lastName']")).sendKeys(lastName);
//            actions.sendKeys(Keys.BACK_SPACE).perform();
            webDriver.findElement(By.xpath("//input[@name='age']")).sendKeys(Keys.BACK_SPACE);
            webDriver.findElement(By.xpath("//input[@name='age']")).sendKeys(String.valueOf(age));
            webDriver.findElement(By.xpath("//input[@name=\"email\"]")).click();
            webDriver.findElement(By.xpath("//input[@name=\"email\"]")).sendKeys(emailId);
            webDriver.findElement(By.xpath("//input[@name='mobileNumber']")).click();
            webDriver.findElement(By.xpath("//input[@name='mobileNumber']")).sendKeys(String.valueOf(mobile));
            webDriver.findElement(By.xpath("//input[@name='contactNumber']")).click();
            webDriver.findElement(By.xpath("//input[@name='contactNumber']")).sendKeys(String.valueOf(contactNo));
//            WebDriverWait wait = new WebDriverWait(webDriver, Duration.ofSeconds(10));

//            webDriver.findElement(By.xpath("//div[contains(text(),'Marital Status')]")).clear();

            webDriver.findElement(By.xpath("//div[contains(text(),'Marital Status')]")).click();
            actions.sendKeys(Keys.ENTER).perform();

//            webDriver.findElement(By.xpath("//div[contains(text(),'Blood Group')]")).clear();

            webDriver.findElement(By.xpath("//div[contains(text(),'Blood Group')]")).click();
            actions.sendKeys(Keys.ENTER).build().perform();
            Thread.sleep(1000);
//            webDriver.findElement(By.xpath("//div[contains(text(),'Identification Document')]")).click();

            //documents
            webDriver.findElement(By.xpath("//div[contains(text(),'Identification Document')]")).click();
            actions.sendKeys(Keys.ENTER).perform();

            //id no
//            webDriver.findElement(By.xpath("//input[@name='identificationDocumentNumber']")).click();
//            actions.sendKeys(Keys.ENTER).build().perform();
            webDriver.findElement(By.xpath("//input[@name='identificationDocumentNumber']")).sendKeys(panNumber);

            //AddressDetails
            //house no
            webDriver.findElement(By.xpath( "//input[@name='houseFlatNumber']")).sendKeys(houseNo);
            actions.sendKeys(Keys.ENTER).perform();

            //street address
            webDriver.findElement(By.xpath("//input[@name='streetAddress']")).sendKeys(streetAddress);
            actions.sendKeys(Keys.ENTER).perform();

            //country
        webDriver.findElement(By.xpath("//div[contains(text(),'Country')]")).click();
        actions.sendKeys(Keys.ARROW_DOWN).build().perform();
        actions.sendKeys(Keys.ARROW_DOWN).build().perform();
        actions.sendKeys(Keys.ARROW_DOWN).build().perform();
        actions.sendKeys(Keys.ARROW_DOWN).build().perform();
        actions.sendKeys(Keys.ARROW_DOWN).build().perform();
        actions.sendKeys(Keys.ENTER).build().perform();

            //state
        webDriver.findElement(By.xpath("//div[contains(text(),'State')]")).click();
        actions.sendKeys(Keys.ARROW_DOWN).build().perform();
        actions.sendKeys(Keys.ARROW_DOWN).build().perform();
        actions.sendKeys(Keys.ARROW_DOWN).build().perform();
        actions.sendKeys(Keys.ARROW_DOWN).build().perform();
        actions.sendKeys(Keys.ARROW_DOWN).build().perform();
        actions.sendKeys(Keys.ARROW_DOWN).build().perform();
//        actions.sendKeys(Keys.ARROW_DOWN).build().perform();
        actions.sendKeys(Keys.ENTER).build().perform();


        //district;oi.xssf.usermode

            webDriver.findElement(By.xpath("//div[contains(text(),'District')]")).click();

            actions.sendKeys("Pune").build().perform();
            actions.sendKeys(Keys.ENTER).build().perform();

            //pincode
//            Thread.sleep(2000);
//            webDriver.findElement(By.xpath("//input[@id=\"react-select-14-input\"]")).sendKeys("411046");
//            actions.sendKeys(Keys.ENTER).build().perform();
            //pincode
            WebDriverWait wait = new WebDriverWait(webDriver, Duration.ofSeconds(5));
//            wait.pollingEvery(Duration.ofMillis(1000));
//            WebElement pinCodeElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='react-select-14-input']")));
            //            // Check if it's the second iteration
            By dynamicXPath = By.xpath("//*[@id='react-select-" + pinCodeXpathValue +"-input']");
            WebElement pincodeElement = wait.until(ExpectedConditions.visibilityOfElementLocated(dynamicXPath));
            pincodeElement.sendKeys("411046");
            actions.sendKeys(Keys.ENTER).build().perform();
//            Thread.sleep(2000);
            //referal type
            webDriver.findElement(By.xpath("//div[contains(text(),'Referral Type')]")).click();
            actions.sendKeys(Keys.ENTER).perform();
//            actions.sendKeys(Keys.ENTER).perform();

            //referal doctor
            webDriver.findElement(By.xpath("//div[contains(text(),'Referral Doctor')]")).click();
            actions.sendKeys(Keys.ENTER).perform();

//            Thread.sleep(2000);
            //reference number
//            webDriver.findElement(By.xpath("//label[@id=\":r49:-label\"]")).sendKeys(String.valueOf(referenceNumber));

            //visit datails patient source
            webDriver.findElement(By.xpath("//div[contains(text(),'Patient Source *')]")).click();
            actions.sendKeys(Keys.ENTER).build().perform();

            //visit type
            webDriver.findElement(By.xpath("//div[contains(text(),'Visit Type *')]")).click();
            actions.sendKeys(Keys.ENTER).build().perform();

            //patient category
            webDriver.findElement((By.xpath("//div[contains(text(),'Patient Category *')]"))).click();
            actions.sendKeys(Keys.ENTER).build().perform();

            //department
            Thread.sleep(2000);
            webDriver.findElement(By.xpath("(//div[contains(text(),'Department *')])")).click();
            actions.sendKeys("cardiology");
            actions.sendKeys(Keys.ENTER).build().perform();

//            actions.sendKeys(Keys.ARROW_DOWN).build().perform();
//            actions.sendKeys(Keys.ARROW_DOWN).build().perform();
//            actions.sendKeys(Keys.ARROW_DOWN).build().perform();
//            actions.sendKeys(Keys.ARROW_DOWN).build().perform();
//            actions.sendKeys(Keys.ARROW_DOWN).build().perform();
//            actions.sendKeys(Keys.ARROW_DOWN).build().perform();
//            actions.sendKeys(Keys.ARROW_DOWN).build().perform();
//            actions.sendKeys(Keys.ARROW_DOWN).build().perform();
            actions.sendKeys(Keys.ENTER).perform();

            //doctor
            webDriver.findElement(By.xpath("//div[contains(text(),'Doctor *')]")).click();
//            actions.sendKeys("Satish Mane ");
            actions.sendKeys(Keys.ENTER).build().perform();

//            //tarrif
//            webDriver.findElement(By.xpath("//div[contains(text(),'Tariff *')]")).click();


            //Tarrif dispensary
            webDriver.findElement(By.xpath("//div[contains(text(),'Tariff Dispensary')]")).click();
            actions.sendKeys(Keys.ENTER).build().perform();

            //complaint reason
            webDriver.findElement(By.xpath("//input[@name=\"complaints\"]")).sendKeys("Normal Visit");
            actions.build().perform();

            //submit
            webDriver.findElement(By.xpath("//button[normalize-space()='Submit']")).click();

            //print case paper
            webDriver.findElement(By.xpath("//input[@aria-label='controlled']")).click();
            Thread.sleep(1000);

            //proceed
            webDriver.findElement(By.xpath("//button[normalize-space()='Proceed']")).click();

            Thread.sleep(2000);
            // Close workbook and WebDriver
//            workbook.close();
//            webDriver.quit();

            //add new service
            webDriver.findElement(By.xpath("//div[@class=' css-ldxf66']")).click();
            actions.sendKeys("Electrocardiography");
            actions.sendKeys(Keys.ENTER).build().perform();
            Thread.sleep(1000);


            //adding doctor to service
            webDriver.findElement(By.xpath("//div[contains(text(),'Doctors')]")).click();
            actions.sendKeys("Satish Mane");
            actions.sendKeys(Keys.ENTER).build().perform();

            //add new service

            webDriver.findElement(By.xpath("//div[@class=' css-ldxf66']")).click();
//            actions.sendKeys("Blood Sugar- Fasting & PP Estimation");
//            actions.sendKeys(Keys.ENTER).build().perform();
//            Thread.sleep(1000);


            //adding doctor to service
//            webDriver.findElement(By.xpath("//*[@id=\"react-select-32-input\"]")).click();
//            actions.sendKeys("Satish Mane");
//            actions.sendKeys(Keys.ENTER).build().perform();



            //after adding service payment will perform
            webDriver.findElement(By.xpath("//button[normalize-space()='Pay Now']")).click();

            //after click pay credit authorized by
            webDriver.findElement(By.xpath("//div[contains(text(),'Credit Authorized By')]")).click();
            actions.sendKeys(Keys.ENTER).build().perform();

            //after taking authorizesation person click pay now
            webDriver.findElement(By.xpath("//button[contains(@type,'submit')][normalize-space()='Pay Now']")).click();

            //after pay now Generate bill and print
            webDriver.findElement(By.xpath("//button[normalize-space()='Generate Bill And Print']")).click();

            Thread.sleep(6000);
            //after generating bill and print close the print
            webDriver.findElement(By.xpath("//button[normalize-space()='Close']")).click();
//            Thread.sleep(2000);
//
//            System.out.println("getCurrentUrl :"+webDriver.getCurrentUrl());
//            Thread.sleep(3000);
//            webDriver.findElement(By.xpath("//p[normalize-space()='Registration']")).click();
            webDriver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
            webDriver.findElement(By.xpath("//p[normalize-space()='Quick Registration']")).click();
            actions.sendKeys(Keys.ENTER).build().perform();
            webDriver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div/div[1]/div/div/ul/div/div/div/div/li[2]/div/div[2]/a/div/span/p")).click();
//            System.out.println("getCurrentUrl :"+webDriver.getCurrentUrl());
//            webDriver.manage().timeouts().pageLoadTimeout(40, TimeUnit.SECONDS);
            webDriver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
            pinCodeXpathValue +=33;
     }
        webDriver.close();
        System.out.println("All Pateints visted succesfully!!!!");

    }
}


