package com.vtiger.ObjectRepository;


	import org.openqa.selenium.Keys;
	import org.openqa.selenium.WebDriver;
	import org.openqa.selenium.WebElement;
	import org.openqa.selenium.support.FindBy;
	import org.openqa.selenium.support.PageFactory;

	public class CreateNewDocumentPage {
	

		@FindBy(name = "notes_title")
		private WebElement documentTitleTxt;

		@FindBy(xpath ="//body[@class='cke_show_borders']")
		private WebElement documentTextarea;

		@FindBy(xpath ="//a[@title='Bold']")
		private WebElement textBoldTollbar;

		@FindBy(xpath ="//a[@title='Italic']")
		private WebElement textItalicTollbar;

		@FindBy(xpath ="//input[@type='file']")
		private WebElement documentPagefileBtn;
		
		@FindBy(xpath ="//b[.='File Information']/../../following-sibling::tr[4]//input[@title='Save [Alt+S]']")
		private WebElement documentpageSaveBtn;
		
		public CreateNewDocumentPage(WebDriver driver) 
		{
			PageFactory.initElements(driver, this);	
		}

		public void documentPageSendTitle(String title) 
		{
			documentTitleTxt.sendKeys(title);
		}

		public void documentPageSendTextAndCopyAll(String textarea) throws InterruptedException 
		{
			documentTextarea.sendKeys(textarea);
			documentTextarea.sendKeys(Keys.CONTROL+"a");
		}

		public void documentPageBoldAndItalicText() throws InterruptedException 
		{
			textBoldTollbar.click();
			Thread.sleep(2000);
			textItalicTollbar.click();
		}
		
		public void documentPageFileUpload(String filePath) throws InterruptedException 
		{
			 documentPagefileBtn.sendKeys(filePath);
			 Thread.sleep(2000);
			documentpageSaveBtn.click();
			Thread.sleep(2000);
			
		}
		

		
	}


