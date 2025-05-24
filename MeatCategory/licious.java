package MeatCategory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementClickInterceptedException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class licious {

    static WebDriver driver;
    static WebDriverWait wait;
    static int maxRetries = 3;

    public static void main(String[] args) throws InterruptedException, IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Licious Chicken Products");

        // Header Row
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("MRP");
        header.createCell(2).setCellValue("SP");
        header.createCell(3).setCellValue("UOM");
        header.createCell(4).setCellValue("Availability");
        header.createCell(5).setCellValue("Description");
        header.createCell(6).setCellValue("Offer");
        header.createCell(7).setCellValue("URL");

        ChromeOptions options = new ChromeOptions();
        driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        wait = new WebDriverWait(driver, Duration.ofSeconds(50));

        driver.get("https://www.licious.in/");

        Thread.sleep(3000);
        if (!clickCategoryDropdown() || !clickChickenCategory() || !waitForProductsToLoad()) {
            System.out.println("Initial setup failed. Exiting.");
            driver.quit();
            workbook.close();
            return;
        }

        int index = 0;
        int excelRowNum = 1;
        Thread.sleep(3000);
        while (true) {
            try {
               
                boolean isNoResultsError = false;
                int refreshRetries = 0;
                int maxRefreshRetries = 2;

                while (refreshRetries <= maxRefreshRetries) {
                    try {
                        isNoResultsError = driver.findElement(By.xpath("//p[contains(text(),'Oops, no results found')]")).isDisplayed();
                    } catch (Exception ignored) {}

                    if (isNoResultsError) {
                        if (refreshRetries < maxRefreshRetries) {
                            System.out.println("No results found error detected. Attempting to refresh page (Attempt " + (refreshRetries + 1) + ").");
                            driver.navigate().refresh();
                            Thread.sleep(2000);

                            // Check for and handle location prompt after refresh
                            try {
                                WebElement cancelButton = wait.until(ExpectedConditions.elementToBeClickable(
                                        By.xpath("//button[contains(text(),'Cancel')] | //span[contains(text(),'Cancel')] | //button[@aria-label='Close']")));
                                cancelButton.click();
                                System.out.println("Location prompt cancelled.");
                                Thread.sleep(1000);
                            } catch (Exception e) {
                                System.out.println("No location prompt found or unable to cancel.");
                            }

                            refreshRetries++;
                            continue;
                        } else {
                            System.out.println("No results error persists after " + maxRefreshRetries + " refresh attempts. Exiting category.");
                            break;
                        }
                    } else {
                        break;
                    }
                }

                if (isNoResultsError && refreshRetries >= maxRefreshRetries) {
                    break; 
                }

                List<WebElement> productList = driver.findElements(By.xpath("//ul[@class='ProductList_product_list__LJ2AW']/li"));

                if (productList.isEmpty() || index >= productList.size()) {
                    break;
                }

                WebElement product = productList.get(index);
                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block: 'center'});", product);
                Thread.sleep(500);

                try {
                    wait.until(ExpectedConditions.elementToBeClickable(product));
                    product.click();
                } catch (ElementClickInterceptedException e) {
                    ((JavascriptExecutor) driver).executeScript("arguments[0].click();", product);
                }

                
                boolean isErrorPage = false;
                refreshRetries = 0;

                while (refreshRetries <= maxRefreshRetries) {
                    try {
                        isErrorPage = driver.findElement(By.xpath("//p[contains(text(),'Oops! The page you')]")).isDisplayed();
                    } catch (Exception ignored) {
                        try {
                            isErrorPage = driver.findElement(By.xpath("//h1[contains(text(),'Page Not Found')]")).isDisplayed();
                        } catch (Exception ignoredAgain) {}
                    }

                    if (isErrorPage) {
                        if (refreshRetries < maxRefreshRetries) {
                            System.out.println("Error page detected. Attempting to refresh page (Attempt " + (refreshRetries + 1) + ").");
                            driver.navigate().refresh();
                            Thread.sleep(2000);

                           
                            try {
                                WebElement cancelButton = wait.until(ExpectedConditions.elementToBeClickable(
                                        By.xpath("//button[contains(text(),'Cancel')] | //span[contains(text(),'Cancel')] | //button[@aria-label='Close']")));
                                cancelButton.click();
                                System.out.println("Location prompt cancelled.");
                                Thread.sleep(1000);
                            } catch (Exception e) {
                                System.out.println("No location prompt found or unable to cancel.");
                            }

                            refreshRetries++;
                            continue;
                        } else {
                            System.out.println("Error page persists after " + maxRefreshRetries + " refresh attempts. Navigating back.");
                            driver.navigate().back();
                            wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//ul[@class='ProductList_product_list__LJ2AW']/li")));
                            index++;
                            break;
                        }
                    } else {
                        break;
                    }
                }

                if (isErrorPage && refreshRetries >= maxRefreshRetries) {
                    continue;
                }

                Thread.sleep(1500);

                List<WebElement> notAvailableElems = driver.findElements(By.xpath("//*[contains(text(),'Not Available')]"));
                if (!notAvailableElems.isEmpty()) {
                    driver.navigate().back();
                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//ul[@class='ProductList_product_list__LJ2AW']/li")));
                    index++;
                    continue;
                }
                Thread.sleep(3000);

                
                //name 
                String name = "NA";
                try {
                    name = driver.findElement(By.xpath("//h1[@class='title_2 ProductDescription_productName__dABoC']")).getText();
                } catch (Exception e) {
                    try {
                        name = driver.findElement(By.xpath("//h1")).getText();
                    } catch (Exception ignored) {
                    	
                    	
                    }
                }

                String mrp = "NA";
                String sp = "NA";

                try {
                    mrp = driver.findElement(By.xpath("(//span[@class='ProductDescription_pricingContainer_upperPart_priceSection_basePrice__GS3UA'])[1]"))
                            .getText().replace("₹", "").trim();
                } catch (Exception ignored) {
                	
                }
                
//sp
                try {
                    sp = driver.findElement(By.xpath("(//span[@class='ProductDescription_pricingContainer_upperPart_priceSection_price__pneu5'])[1]"))
                            .getText().replace("₹", "").trim();
                } catch (Exception ignored) {}

              
                if (mrp.isEmpty() || mrp.equalsIgnoreCase("NA")) {
                    mrp = sp;
                }

                String uom = "NA";
                try {
                    uom = driver.findElement(By.xpath("//div[@class='QuantityTable_qtyBlock__54Ov5']")).getText();
                } catch (Exception ignored) {}

                String availability = "0";
                try {
                    WebElement btn = driver.findElement(By.xpath("//div[contains(@class, 'AddToCartButtonPLP_add_btn__')]"));
                    String btnText = btn.getText().toLowerCase();

                    if (btnText.contains("add")) {
                        availability = "1";
                    } else if (btnText.contains("notify")) {
                        availability = "0";
                    }
                } catch (Exception e) {
                    availability = "0";
                }


                Thread.sleep(8000);
                try {
                    List<WebElement> readMoreElements = driver.findElements(By.xpath("//span[contains(text(),'Read More')]"));
                    if (!readMoreElements.isEmpty()) {
                        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", readMoreElements.get(0));
                        readMoreElements.get(0).click();
                        Thread.sleep(500);
                    }
                } catch (Exception e) {
                    System.out.println("Read More not clickable or not needed.");
                }

                String des = "NA";
                try {
                    WebElement desElem = driver.findElement(By.xpath("//div[@class='ProductDescription_productDescription_detail__vAc5H']"));
                    des = desElem.getText().trim();
                } catch (Exception e) {
                    System.out.println("Description not found.");
                }

                String offer = "NA";
                try {
                    WebElement offerElem = driver.findElement(By.xpath("(//span[@class='ProductDescription_pricingContainer_upperPart_priceSection_discount__MgDSj'])[1]"));
                    offer = offerElem.getText().trim();
                } catch (Exception e) {
                    System.out.println("Offer not found.");
                }

                String url = "";
                try {
                    url = driver.getCurrentUrl();
                } catch (Exception e) {
                    System.out.println("URL not found.");
                }

                System.out.println("Product :" + (index + 1));
                System.out.println("Name: " + name);
                System.out.println("MRP: " + mrp);
                System.out.println("SP: " + sp);
                System.out.println("UOM: " + uom);
                System.out.println("Availability: " + availability);
                System.out.println("Description: " + des);
                System.out.println("Offer: " + offer);
                System.out.println("URL: " + url);
                System.out.println("--------------------------------------------------");

                Row row = sheet.createRow(excelRowNum++);
                row.createCell(0).setCellValue(name);
                row.createCell(1).setCellValue(mrp);
                row.createCell(2).setCellValue(sp);
                row.createCell(3).setCellValue(uom);
                row.createCell(4).setCellValue(availability);
                row.createCell(5).setCellValue(des);
                row.createCell(6).setCellValue(offer);
                row.createCell(7).setCellValue(url);

                driver.navigate().back();
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//ul[@class='ProductList_product_list__LJ2AW']/li")));
                index++;

            } catch (StaleElementReferenceException e) {
                System.out.println("Stale element. Retrying category.");
                if (!clickCategoryDropdown() || !clickChickenCategory() || !waitForProductsToLoad()) {
                    break;
                }

            } catch (Exception e) {
                System.out.println("Error at product #" + (index + 1) + ": " + e.getMessage());
                index++;
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//ul[@class='ProductList_product_list__LJ2AW']/li")));
            }
        }

        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        String filePath = ".//Output_licious_Chicken" + timestamp + ".xlsx";

        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
            System.out.println("Data saved to: " + filePath);
        } catch (IOException e) {
            System.out.println("Failed to save Excel: " + e.getMessage());
        }

        workbook.close();
        driver.quit();
    }

    private static boolean clickCategoryDropdown() throws InterruptedException {
        int retries = 0;
        while (retries < maxRetries) {
            try {
                WebElement category = wait.until(ExpectedConditions.elementToBeClickable(
                        By.xpath("//div[@class='CategoryDropdown_categories_button__Z1XVX']//span")));
                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", category);
                category.click();
                return true;
            } catch (Exception e) {
                retries++;
                Thread.sleep(1000);
            }
        }
        return false;
    }

    private static boolean clickChickenCategory() throws InterruptedException {
        int retries = 0;
        while (retries < maxRetries) {
            try {
                WebElement chickenCategory = wait.until(ExpectedConditions.elementToBeClickable(
                        By.xpath("(//ul[@class='CategoryMenu_category_list__czV2_']//li)[2]")));
                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", chickenCategory);
                chickenCategory.click();
                return true;
            } catch (Exception e) {
                retries++;
                Thread.sleep(1000);
            }
        }
        return false;
    }

    private static boolean waitForProductsToLoad() throws InterruptedException {
        int retries = 0;
        while (retries < maxRetries) {
            try {
                wait.until(ExpectedConditions.visibilityOfElementLocated(
                        By.xpath("//ul[@class='ProductList_product_list__LJ2AW']/li")));
                return true;
            } catch (Exception e) {
                retries++;
                Thread.sleep(1000);
            }
        }
        return false;
    }
}