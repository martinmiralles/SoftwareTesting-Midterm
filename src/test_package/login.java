package test_package;
import org.openqa.selenium.WebDriver;
//comment the above line and uncomment below line to use Chrome
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.*;
import org.openqa.selenium.By;
//comment the above line and uncomment below line to use Chrome

public class login {

    public static void main(String[] args) {
        // declaration and instantiation of objects/variables
//        System.setProperty("webdriver.gecko.driver","C:\\Users\\Mar\\Downloads\\chromedriver_win32\\chromedriver.exe");
//        WebDriver driver = new FirefoxDriver();
        //comment the above 2 lines and uncomment below 2 lines to use Chrome
        System.setProperty("webdriver.chrome.driver","C:\\Users\\300965925\\Downloads\\chromedriver_win32\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        String baseUrl = "http://demo.guru99.com/test/login.html";
        driver.get(baseUrl);

        // Get the WebElement corresponding to the Email Address(TextField)
        WebElement email = driver.findElement(By.id("email"));

        // Get the WebElement corresponding to the Password Field
        WebElement password = driver.findElement(By.name("passwd"));

        email.sendKeys("abcd@gmail.com");
        password.sendKeys("abcdefghlkjl");
        System.out.println("Text Field Set");

        // Deleting values in the text box
        email.clear();
        password.clear();
        System.out.println("Text Field Cleared");

        // Find the submit button
        WebElement login = driver.findElement(By.id("SubmitLogin"));

        // Using click method to submit form
        email.sendKeys("abcd@gmail.com");
        password.sendKeys("abcdefghlkjl");
        login.click();
        System.out.println("Login Done with Click");

        //using submit method to submit the form. Submit used on password field
        driver.get(baseUrl);
        driver.findElement(By.id("email")).sendKeys("abcd@gmail.com");
        driver.findElement(By.name("passwd")).sendKeys("abcdefghlkjl");
        driver.findElement(By.id("SubmitLogin")).submit();
        System.out.println("Login Done with Submit");

        //driver.close();

    }

}
