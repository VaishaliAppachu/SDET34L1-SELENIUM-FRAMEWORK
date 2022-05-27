package com.vtiger.ObjectRepository;

	import org.openqa.selenium.WebDriver;
	import org.openqa.selenium.WebElement;
	import org.openqa.selenium.support.FindBy;
	import org.openqa.selenium.support.PageFactory;
	// create class 
	public class LoginPage {
		@FindBy(xpath = "//input[@type='text']")
		private WebElement userNameTxt;
		@FindBy(xpath = "//input[@type='password']")
		private WebElement passwordTxt;
		@FindBy(xpath = "//input[@id='submitButton']")
		private WebElement SigninBtn;
		
		public LoginPage(WebDriver driver) {
		PageFactory.initElements(driver,this);
		}
		
		public void loginAction(String userName,String password) {
			userNameTxt.sendKeys(userName);
			passwordTxt.sendKeys(password);
			SigninBtn.click();
		}
		
}
