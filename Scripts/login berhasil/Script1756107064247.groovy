import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSSFWorkbook
import org.apache.poi.ss.usermodel.*
import java.io.FileOutputStream as FileOutputStream
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Cell


//buka web
WebUI.openBrowser('')

WebUI.maximizeWindow()

WebUI.navigateToUrl('https://www.saucedemo.com/')

WebUI.delay(2)

String ss1 = 'Screenshots/login_page.png'

WebUI.takeScreenshot(ss1)

// input username
WebUI.setText(findTestObject('Object Repository/Page_Swag Labs/input_Swag Labs_user-name'), 'standard_user')

// input password
WebUI.setText(findTestObject('Object Repository/Page_Swag Labs/input_Swag Labs_password'), 'secret_sauce')

WebUI.delay(2)

String ss2 = 'Screenshots/after_input.png'

WebUI.takeScreenshot(ss2)

// klik login
WebUI.click(findTestObject('Object Repository/Page_Swag Labs/input_Swag Labs_login-button'))

WebUI.delay(3)

String ss3 = 'Screenshots/main_menu.png'

WebUI.takeScreenshot(ss3)

// ==== STEP 4: Simpan hasil ke Excel ====
String filePath = 'C:/Users/MII/Katalon Studio/web testing/Reports/LoginResult.xlsx'

// Buat workbook & sheet
XSSFWorkbook workbook = new XSSFWorkbook()

Sheet sheet = workbook.createSheet('Login Test')

// Header
Row header = sheet.createRow(0)

header.createCell(0).setCellValue('Step')

header.createCell(1).setCellValue('Screenshot Path')

// Baris 1
Row row1 = sheet.createRow(1)

row1.createCell(0).setCellValue('Open Website (Login Page)')

row1.createCell(1).setCellValue(ss1)

// Baris 2
Row row2 = sheet.createRow(2)

row2.createCell(0).setCellValue('After Input Username & Password')

row2.createCell(1).setCellValue(ss2)

// Baris 3
Row row3 = sheet.createRow(3)

row3.createCell(0).setCellValue('After Login Success (Main Menu)')

row3.createCell(1).setCellValue(ss3)

// Simpan file Excel
FileOutputStream fileOut = new FileOutputStream(filePath)

workbook.write(fileOut)

fileOut.close()

workbook.close()

// ==== STEP 5: Tutup browser ====
WebUI.closeBrowser()

