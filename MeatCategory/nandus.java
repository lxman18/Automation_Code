package MeatCategory;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.*;

public class nandus {
    public static void main(String[] args) throws InterruptedException {
        ChromeOptions options = new ChromeOptions();
        WebDriver driver = new ChromeDriver(options);
        JavascriptExecutor js = (JavascriptExecutor) driver;
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));
        driver.manage().window().maximize();
        
        
        String mrp = "";
        String sp = "";
        String uom ="";
        String availability = "0";
        

        // Create Excel Workbook and Sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Product Data");

        // Create header row
        Row headerRow = sheet.createRow(0);
        String[] headers = {"Name", "MRP", "SP", "UOM", "Availability", "URL"};
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }

        int rowNum = 1;

        List<String> urls = Arrays.asList(
            "https://www.nandus.com/product/chicken",
            "https://www.nandus.com/product/fish-and-seafood-1",
            "https://www.nandus.com/product/eggs"
        );

        By productCardBy = By.xpath("//div[@class='productLists ng-star-inserted']//div[contains(@class,'card')]/div");

        for (String url : urls) {
            if (!url.startsWith("http")) {
                System.out.println("Skipping invalid URL: " + url);
                continue;
            }

            driver.get(url);
            wait.until(ExpectedConditions.visibilityOfElementLocated(productCardBy));
            List<WebElement> productCards = driver.findElements(productCardBy);
            int productCount = 1;
            for (int i = 0; i < productCards.size(); i++) {
                try {
                	
                	
                    productCards = driver.findElements(productCardBy);
                    WebElement product = productCards.get(i);
                    js.executeScript("arguments[0].scrollIntoView({block: 'center'});", product);
                    Thread.sleep(500);
                    js.executeScript("arguments[0].click();", product);

                    Thread.sleep(4000);
                    wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.xpath("//div[@class='col-md-6 col-sm-6 col-xs-12 ui_v2 mobile-header hidden-xs']//h1")));

                    String name = safeGetText(driver, By.xpath("//div[@class='col-md-6 col-sm-6 col-xs-12 ui_v2 mobile-header hidden-xs']//h1"));
                    if (name.isEmpty()) name = safeGetText(driver, By.xpath("//h1"));

                    

                    Thread.sleep(3000);

                    try {
                        mrp = safeGetText(driver, By.xpath("//del[@class='del_price txt-brown h5_medium']//div"))
                                  .replace("Rs.", "").trim();
                    } catch (Exception e) {
                        mrp = ""; 
                    }
                    Thread.sleep(3000);
                    try {
                        sp = safeGetText(driver, By.xpath("//div[@class='d-flex justify-content-start']//p"))
                                .replace("Rs.", "").trim();
                    } catch (Exception e) {
                        try {
							sp = safeGetText(driver, By.xpath("//div[@class='combo-product-price']//div[@class='d-flex justify-content-start']//p"))
							        .replace("Rs.", "").trim();
						} catch (Exception e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
                    }

                    
                    if (mrp.isEmpty()) {
                        mrp = sp;
                    }

                   
					try {
						uom = safeGetText(driver, By.xpath("//*[@id=\"product-details\"]/div/div/div[3]/div[3]/div[1]/div[1]/div[2]"));
						
					} catch (Exception e) {
						if (uom.isEmpty()) uom = "NA";
					}

                    
                    try {
                        WebElement addBtn = driver.findElement(By.xpath("(//button[@class='btn btn-red ng-star-inserted'])[1]"));
                        availability = (addBtn.isDisplayed() && addBtn.isEnabled()) ? "1" : "0";
                    } catch (Exception e) {
                        availability = "0";
                    }

                    String productUrl = driver.getCurrentUrl();
                    System.out.println("------------Product Count:"+ productCount++ +"---------------------------");
                    System.out.println("Product Name: " + name);
                    System.out.println("MRP: " + mrp);
                    System.out.println("SP: " + sp);
                    System.out.println("UOM: " + uom);
                    System.out.println("Availability: " + availability);
                    System.out.println(productUrl);
                    System.out.println("-----------------------");
                    // Write row to Excel
                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(name);
                    row.createCell(1).setCellValue(mrp);
                    row.createCell(2).setCellValue(sp);
                    row.createCell(3).setCellValue(uom);
                    row.createCell(4).setCellValue(availability);
                    row.createCell(5).setCellValue(productUrl);

                    driver.navigate().back();
                    wait.until(ExpectedConditions.visibilityOfElementLocated(productCardBy));
                    Thread.sleep(1000);
                } catch (Exception e) {
                    e.printStackTrace();
                    try {
                        driver.navigate().back();
                        wait.until(ExpectedConditions.visibilityOfElementLocated(productCardBy));
                        Thread.sleep(1000);
                    } catch (Exception ex) {
                        System.out.println("Failed to recover after error.");
                    }
                }
            }
        }

        // Auto-size columns
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }

        
        SimpleDateFormat data = new SimpleDateFormat("dd_MM_YYYY_HH_Mm_ss");
		String timestamp = data.format(new Date() );
        try (FileOutputStream fileOut = new FileOutputStream(".//Output//Nandus_"+timestamp+".xlsx")) {
            workbook.write(fileOut);
            workbook.close();
            System.out.println("Excel file written successfully.");
        } catch (Exception e) {
            e.printStackTrace();
        }


        driver.quit();
    }

    private static String safeGetText(WebDriver driver, By by) {
        try {
            return driver.findElement(by).getText().trim();
        } catch (Exception e) {
            return "";
        }
    }
}
