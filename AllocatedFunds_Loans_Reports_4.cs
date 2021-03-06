'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Constants
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT SWIFT_International_Payorder_Library
Option Explicit

'Test case ID 161860
'Test case ID 161862

Dim folderName, sDATE, fDATE, colName(5), param
Dim limitChanges, calculationDates, availabilityDates, registryInputInfo, numOfExpiredDays, subjectivity
Dim actualFile1, actualFile2, actualFile3, actualFile4, actualFile5, actualFile6
Dim expectedFile1, expectedFile2, expectedFile3, expectedFile4, expectedFile5, expectedFile6
Dim resultFile1, resultFile2, resultFile3, resultFile4, resultFile5, resultFile6

Sub AllocatedFunds_Loans_Reports_4(rowLimit)
		' ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|ì³ñÏ»ñ (î»Õ³µ³ßËí³Í)
		Call Test_Initialize()

		' Համակարգ մուտք գործել ARMSOFT օգտագործողով
		Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
  Call Test_StartUp(rowLimit) 
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''ê³ÑÙ³Ý³ã³÷»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ''''''''''
		
		' Լրացնել Սահմանաչափերի փոփոխություններ դիալոգային պատուհանը
		Log.Message "Սահմանաչափերի փոփոխություններ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "ê³ÑÙ³Ý³ã³÷»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ", limitChanges)
		
		if WaitForExecutionProgress() then		
				' êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 4, "frmPttel")
				' Արտահանել, որպես txt ֆայլ
				Call ExportToTXTFromPttel("frmPttel", actualFile1)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 4)
				' Համեմատել txt ֆայլերը
				Call Compare_Files(actualFile1, expectedFile1, param)
				' ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ³Ùë³Ãí»ñ''''''''''
		
		' Լրացնել Տոկոսների հաշվարկման ամսաթվեր դիալոգային պատուհանը
		Log.Message "Տոկոսների հաշվարկման ամսաթվեր", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ³Ùë³Ãí»ñ", calculationDates)
		
		if WaitForExecutionProgress() then		
				' êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 4, "frmPttel")
				' Արտահանել Excel
				Call ExportToExcel("frmPttel", actualFile2)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 31137)
				' Համեմատել Excel ֆայլերը
				Call CompareTwoExcelFiles(actualFile2, expectedFile2, resultFile2)
				' ö³Ï»É բոլոր Excel ֆայլերը
				Call CloseAllExcelFiles()
				' ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		'''''''''''''îñ³Ù³¹ñáõÙÝ»ñÇ ³Ùë³Ãí»ñ'''''''''''''
		
		' Լրացնել Տրամադրումների ամսաթվեր դիալոգային պատուհանը
		Log.Message "Տրամադրումների ամսաթվեր", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "îñ³Ù³¹ñáõÙÝ»ñÇ ³Ùë³Ãí»ñ", availabilityDates)
		
		if WaitForExecutionProgress() then		
				' Արտահանել, որպես txt ֆայլ
				Call ExportToTXTFromPttel("frmPttel", actualFile3)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 1)
				' Համեմատել txt ֆայլերը
				Call Compare_Files(actualFile3, expectedFile3, param)
				' ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		'''''''è»·ÇëïñÇ Ùáõïù³ÛÇÝ ï»Õ»Ï³ïíáõÃÛáõÝ''''''''
		
		folderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|è»·ÇëïñÇ Ùáõïù³ÛÇÝ ï»Õ»Ï³ïíáõÃÛáõÝ|"
		
		' Լրացնել Ռեգիստրի մուտքային տեղեկատվություն դիալոգային պատուհանը
		Log.Message "Ռեգիստրի մուտքային տեղեկատվություն", "", pmNormal, DivideColor
		Call GoTo_RegistryInputInformation(folderName, "è»·ÇëïñÇ Ùáõïù³ÛÇÝ ï»Õ»Ï³ïíáõÃÛáõÝ", registryInputInfo)
		
		if WaitForExecutionProgress() then		
				' êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 4, "frmPttel")
				' Արտահանել Excel
				Call ExportToExcel("frmPttel", actualFile4)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 224)
				' Համեմատել Excel ֆայլերը
				Call CompareTwoExcelFiles(actualFile4, expectedFile4, resultFile4)
				' ö³Ï»É բոլոր Excel ֆայլերը
				Call CloseAllExcelFiles()
				' ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''Ä³ÙÏ»ï³Ýó ûñ»ñÇ ù³Ý³Ï''''''''''''''
		
		' Լրացնել Ժամկետանց օրերի քանակ դիալոգային պատուհանը
		Log.Message "Ժամկետանց օրերի քանակ", "", pmNormal, DivideColor
		Call GoTo_RegistryInputInformation(folderName & "¶áñÍáÕáõÃÛáõÝÝ»ñÇ ¹ÇïáõÙ|", "Ä³ÙÏ»ï³Ýó ûñ»ñÇ ù³Ý³Ï", numOfExpiredDays)
		
		if WaitForExecutionProgress() then		
				' Արտահանել, որպես txt ֆայլ
				Call ExportToTXTFromPttel("frmPttel", actualFile5)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 2)
				' Համեմատել txt ֆայլերը
				Call Compare_Files(actualFile5, expectedFile5, param)
				' ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''êáõµÛ»ÏïÇíáõÃÛáõÝ''''''''''''''''
		
		' Լրացնել Սուբյեկտիվություն դիալոգային պատուհանը
		Log.Message "Սուբյեկտիվություն", "", pmNormal, DivideColor
		Call GoTo_RegistryInputInformation(folderName  & "¶áñÍáÕáõÃÛáõÝÝ»ñÇ ¹ÇïáõÙ|", "êáõµÛ»ÏïÇíáõÃÛáõÝ", subjectivity)
		
		if WaitForExecutionProgress() then		
				' êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 4, "frmPttel")
				' Արտահանել Excel
				Call ExportToExcel("frmPttel", actualFile6)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 224)
				' Համեմատել Excel ֆայլերը
				Call CompareTwoExcelFiles(actualFile6, expectedFile6, resultFile6)
				' ö³Ï»É բոլոր Excel ֆայլերը
				Call CloseAllExcelFiles()
				' ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
				Call Close_AsBank()		
End	Sub

Sub Test_StartUp(rowLimit)
		Call Initialize_AsBank("bank_Report", sDATE, fDATE)
  Login("ARMSOFT")
		Call SaveRAM_RowsLimit(rowLimit)
		Call ChangeWorkspace(c_Subsystems)
End	Sub

Sub Test_Initialize()
		folderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|¶áñÍáÕáõÃÛáõÝÝ»ñ, ÷á÷áËáõÃÛáõÝÝ»ñ|²ÛÉ|"
	
		sDATE = "20030101"
		fDATE = "20260101"  
		
		colName(0) = "fKEY"
		colName(3) = "fCOM"
		colName(1) = "fDATE"
		colName(2) = "fSUID"
		
		' ê³ÑÙ³Ý³ã³÷»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		expectedFile1 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Expected\expectedFile1.txt"
		' îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ³Ùë³Ãí»ñ
		expectedFile2 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Expected\expectedFile2.xlsx"
		' îñ³Ù³¹ñáõÙÝ»ñÇ ³Ùë³Ãí»ñ
		expectedFile3 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Expected\expectedFile3.txt"
		' è»·ÇëïñÇ Ùáõïù³ÛÇÝ ï»Õ»Ï³ïíáõÃÛáõÝ
		expectedFile4 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Expected\expectedFile4.xlsx"
		' Ä³ÙÏ»ï³Ýó ûñ»ñÇ ù³Ý³Ï
		expectedFile5 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Expected\expectedFile5.txt"
		' êáõµÛ»ÏïÇíáõÃÛáõÝ
		expectedFile6 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Expected\expectedFile6.xlsx"
	
  ' ê³ÑÙ³Ý³ã³÷»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		actualFile1 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Actual\actualFile1.txt"
		' îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ³Ùë³Ãí»ñ
		actualFile2 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Actual\actualFile2.xlsx"
		' îñ³Ù³¹ñáõÙÝ»ñÇ ³Ùë³Ãí»ñ
		actualFile3 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Actual\actualFile3.txt"
		' è»·ÇëïñÇ Ùáõïù³ÛÇÝ ï»Õ»Ï³ïíáõÃÛáõÝ
		actualFile4 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Actual\actualFile4.xlsx"
		' Ä³ÙÏ»ï³Ýó ûñ»ñÇ ù³Ý³Ï
		actualFile5 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Actual\actualFile5.txt"
		' êáõµÛ»ÏïÇíáõÃÛáõÝ
		actualFile6 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Actual\actualFile6.xlsx"
		
  ' îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ³Ùë³Ãí»ñ
		resultFile2 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Result\resultFile2.xlsx"
		' è»·ÇëïñÇ Ùáõïù³ÛÇÝ ï»Õ»Ï³ïíáõÃÛáõÝ
		resultFile4 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Result\resultFile4.xlsx"
		' êáõµÛ»ÏïÇíáõÃÛáõÝ
		resultFile6 = Project.Path & "Stores\Reports\Subsystems\Allocated Funds\LoansTest4\Result\resultFile6.xlsx"
		
  ' ê³ÑÙ³Ý³ã³÷»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		Set limitChanges = New_AgreementsCommomFilter()
		with limitChanges
				.note2 = "01"
				.agreeOffice = "P04"
				.onlyChangesExists = true
		end with
		
		' îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ³Ùë³Ãí»ñ
		Set calculationDates = New_AgreementsCommomFilter()
		with calculationDates
				.startDate = "27/10/14"
		end with
		
		' îñ³Ù³¹ñáõÙÝ»ñÇ ³Ùë³Ãí»ñ
		Set availabilityDates = New_AgreementsCommomFilter()
		with availabilityDates
				.startDate = "13/07/17"
				.endDate = "13/07/17"
				.agreeN = "TV22127"
				.performer = "253"
				.agreeOffice = "P00"
				.agreeSection = "08"
				.accessType = "C11"
		end with

		' è»·ÇëïñÇ Ùáõïù³ÛÇÝ ï»Õ»Ï³ïíáõÃÛáõÝ		
		Set registryInputInfo = New_RegistryInputInformation()
		with registryInputInfo
				.endDate = "01/01/23"
				.NumOfexpDaysExists = true
		end with
		
		' Ä³ÙÏ»ï³Ýó ûñ»ñÇ ù³Ý³Ï
		Set numOfExpiredDays = New_RegistryInputInformation()
		with numOfExpiredDays
				.startDate = "16/01/14"
    .endDate = "10/02/14"
    .client = "00023337"
				.clientDataExists = false
				.clientData = 1
		end with
		
		' êáõµÛ»ÏïÇíáõÃÛáõÝ
		Set subjectivity = New_RegistryInputInformation()
		with subjectivity
				.endDate = "01/01/23"
				.clientDataExists = false
		end with
End Sub