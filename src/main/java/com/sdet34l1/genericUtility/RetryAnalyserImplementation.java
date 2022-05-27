package com.sdet34l1.genericUtility;

import org.testng.IRetryAnalyzer;
import org.testng.ITestResult;
/**
 * This method is used to rerun the failed TestScripts
 * @author VAISHALI
 * 
 */
public class RetryAnalyserImplementation implements IRetryAnalyzer{
	int count=0;
	int maxRetry=4;
	@Override
	public boolean retry(ITestResult result) {
		
		while(count<maxRetry)
		{
			count++;
			return true;
			
		}
		return false;
	}

}
