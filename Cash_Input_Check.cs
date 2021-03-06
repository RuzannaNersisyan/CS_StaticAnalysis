'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Constants
'USEUNIT Library_Colour
'USEUNIT DAHK_Library_Filter
'USEUNIT CashInput_Confirmphases_Library
'USEUNIT Payment_Except_Library
'USEUNIT Akreditiv_Library
'USEUNIT Main_Accountant_Filter_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_CheckDB 
Option Explicit

'Test Case ID 177555

Dim sDate, eDate, folderName, expectedFile, actualFilePath, actualFile
Dim cashInputCreate, cashInputEdit, workingDocs, verifyDoc, currentDate, param
Dim fBODY, dbo_FOLDERS(2)

Sub Cash_Input_Check_Test()
				Call Test_Inintialize()

				' Համակարգ մուտք գործել ARMSOFT օգտագործողով
				Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
				Call Test_StartUp()
				
				' Մուտք գործել Հաշիվներ թղթապանակ
				Log.Message "Մուտք գործել Հաշիվներ թղթապանակ", "", pmNormal, DivideColor
				Call OpenAccauntsFolder(folderName & "Ð³ßÇíÝ»ñ","1","","33170080101","","","","","","",0,"","","","","",0,0,0,"","","","","","ACCS","0")		
				Call CheckPttel_RowCount("frmPttel", 1) 
		
				' Ստեղծել Կանխիկ մուտք փաստաթուղթ
				Log.Message "Ստեղծել Կանխիկ մուտք փաստաթուղթ", "", pmNormal, DivideColor
				Call Create_Cash_Input(cashInputCreate, "ê¨³·Çñ")
    
    ' Փակել Հաշիվներ թղթապանակը
				Call Close_Window(wMDIClient, "frmPttel")
    
    ' Կանխիկ մուտք փաստաթղթի ստեղծումից հետո SQL ստուգում
    Log.Message "Կանխիկ մուտք փաստաթղթի ստեղծումից հետո SQL ստուգում", "", pmNormal, DivideColor
    Call DB_Initialize()
    Call Check_DB_Draft()
    
    ' Մուտք գործել Օգտագործողսի Սևագրեր թղթապանակ
    Log.Message "Մուտք գործել Օգտագործողսի Սևագրեր թղթապանակ", "", pmNormal, DivideColor
    Call wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|ú·ï³·áñÍáÕÇ ë¨³·ñ»ñ")
    ' Կատարել կոճակի սեղմում
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    ' Ստեղծել Կանխիկ մուտք փաստաթուղթը սևագրերից
    Log.Message "Ստեղծել Կանխիկ մուտք փաստաթուղթը սևագրերից", "", pmNormal, DivideColor
    If SearchInPttel("frmPttel", 2, cashInputCreate.fIsn) Then
        With cashInputCreate
            .coinTab.coinForCheck = "0.89"
            .commonTab.payer = "00000233"
            .chargeTab.chargeAcc = "000003701  "
            .chargeTab.chargeCurr = "001"
            .chargeTab.chargeAmount = "234.49"
            .chargeTab.chargePercent = "0.5000"
            .chargeTab.incomeAcc = "000527500  "
            .chargeTab.incomeAccCurr = "000"
            .chargeTab.comment = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
        End With
        BuiltIn.Delay(3000)
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_ToEdit)
        If wMDIClient.WaitvbObject("frmASDocForm", 3000).Exists Then
            Call Fill_CashIn_Common(cashInputCreate.commonTab)
            Call Fill_CashIn_Charge(cashInputCreate.chargeTab, cashInputCreate.commonTab.payerLegalStatus)
            Call Check_Cash_Input(cashInputCreate)
            Call ClickCmdButton(1, "Î³ï³ñ»É")
        End If
    Else
        Log.Error "Can't find searched row.", "", pmNormal, ErrorColor    
    End If
    
    ' Քաղվածքի պահպանում 
				Log.Message "Քաղվածքի պահպանում", "", pmNormal, DivideColor
    Call SaveDoc(actualFilePath, actualFile) 
				
				' Փակել Քաղվածքի պատուհանը 
				Call Close_Window(wMDIClient, "FrmSpr")
    
    ' Փակել Օգտագործողի սևագրեր թղթապանակը
				Call Close_Window(wMDIClient, "frmPttel")
    
    ' Կանխիկ մուտք փաստաթուղթը սևագրերից ստեղծումից հետո SQL ստուգում
    Log.Message "Կանխիկ մուտք փաստաթուղթը սևագրերից ստեղծումից հետո SQL ստուգում", "", pmNormal, DivideColor
    Call Check_DB_Create()
				
				' Փաստացի քաղվածքի համեմատում սպասվողի հետ
				Log.Message "Փաստացի քաղվածքի համեմատում սպասվողի հետ", "", pmNormal, DivideColor
				param = "N\s\d{1,6}\s*.\d{1,10}\s{0,}.|Date\s\d{1,2}.\d{1,2}.\d{1,2}\s(\d{1,2}:\d{1,2})*"
    Call Compare_Files(actualFilePath & actualFile, expectedFile, param)
				
				' Մուտք գործել Աշխատանքային փաստաթղթեր թղթապանակ
				folderName = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|"
				Call GoTo_MainAccWorkingDocuments(folderName, workingDocs)
				
				' Ստուգել ստեղծված փաստաթղթի արժեքները և խմբագրել այն 
				Log.Message "Ստուգել ստեղծված փաստաթղթի արժեքները և խմբագրել այն", "", pmNormal, DivideColor
				If SearchInPttel("frmPttel", 2, cashInputCreate.commonTab.docNum) Then
    				Call Edit_Cash_Input(cashInputCreate, cashInputEdit, "Î³ï³ñ»É")
    Else
        Log.Error "Can't find searched row.", "", pmNormal, ErrorColor
    End If
    With cashInputEdit
        .commonTab.payer = ""
        .chargeTab.operAreaForCheck = "9P"
        .chargeTab.legalStatusForCheck = "12"
        .coinTab.coinForCheck = "0.85"
        .chargeTab.commentForCheck = "²ñï³ñÅ.·³ÝÓáõÙ"
    End With
    
    ' Կանխիկ մուտք փաստաթուղթը խմբագրելուց հետո SQL ստուգում
    Log.Message "Կանխիկ մուտք փաստաթուղթը խմբագրելուց հետո SQL ստուգում", "", pmNormal, DivideColor
    Call Check_DB_Edit()
				
				' Ուղարկել հաստատման
				Log.Message "Ուղարկել հաստատման", "", pmNormal, DivideColor
				Call SendToVerify_Contrct(3, 2, "Î³ï³ñ»É")
    
    ' Ուղարկել հաստատման-ից հետո SQL ստուգում
    Log.Message "Ուղարկել հաստատման-ից հետո SQL ստուգում", "", pmNormal, DivideColor
    Call Check_DB_SendToVerify()
				
				' Մուտք գործել Հաստատվող փաստաթղթեր (|) թղթապանակ
				Call GoToVerificationDocument(folderName & "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ (I)", verifyDoc)
				
				' Վավերացնել փաստաթուղթը
				Log.Message "Վավերացնել փաստաթուղթը", "", pmNormal, DivideColor
    If SearchInPttel("frmPttel", 1, cashInputEdit.fIsn) Then
    				Call Validate_Doc()
    Else
        Log.Error "Can't find searched row.", "", pmNormal, ErrorColor
    End If
				
				' Փակել Հաստատվող փաստաթղթեր (|) թղթապանակը
				Call Close_Window(wMDIClient, "frmPttel")
    
    ' Վավերացումից հետո SQL ստուգում
    Log.Message "Վավերացումից հետո SQL ստուգում", "", pmNormal, DivideColor
    Call Check_DB_Validate()
				
				' Մուտք գործել Ստեղծված փաստաթղթեր թղթապանակ
				currentDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d%m%y")
				Call OpenCreatedDocFolder(folderName & "êï»ÕÍí³Í ÷³ëï³ÃÕÃ»ñ", currentDate, currentDate, null, "KasPrOrd")
				
				' Ստուգել, որ առկա է մեր ավելացրած փաստաթուղթը
				Log.Message "Ստուգել, որ առկա է մեր ավելացրած փաստաթուղթը", "", pmNormal, DivideColor
				If SearchInPttel("frmPttel", 2, cashInputEdit.fIsn) Then
    				wMDIClient.vbObject("frmPttel").Keys("^w")
        If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
            Call Check_Cash_Input(cashInputEdit)
            Call ClickCmdButton(1, "OK") 
        Else 
            Log.Error "Can't open frmASDocForm window.", "", pmNormal, ErrorColor
        End If
				Else
        Log.Error "Can't find searched row.", "", pmNormal, ErrorColor
    End If
				
				' Ջնջել Կանխիկի մուտք փաստաթուղթը
				Log.Message "Ջնջել Կանխիկի մուտք փաստաթուղթը", "", pmNormal, DivideColor
				Call SearchAndDelete("frmPttel", 2, cashInputEdit.fIsn, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
				
				' Փակել Ստեղծված փաստաթղթեր թղթապանակը
				Call Close_Window(wMDIClient, "frmPttel")
    
    ' Ջնջելուց հետո SQL ստուգում
    Log.Message "Ջնջելուց հետո SQL ստուգում", "", pmNormal, DivideColor
    Call Check_DB_Delete()
				
				' Փակել ծրագիրը
				Call Close_AsBank()
End	Sub

Sub Test_StartUp()
				Call Initialize_AsBank("bank", sDate, eDate)   
				Login("ARMSOFT")
				' Մուտք Գլխավոր հաշվապահի ԱՇՏ
				Call ChangeWorkspace(c_ChiefAcc)
End Sub

Sub Test_Inintialize()
				sDate = "20030101"
				eDate = "20250101"
		
				folderName = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|"
				expectedFile = Project.Path &  "Stores\Cash_Input_Output\Expected\Expected_Cash_Input_Create.txt"
				actualFilePath = Project.Path &  "Stores\Cash_Input_Output\Actual\"
    actualFile = "Actual_Cash_Input_Create.txt"
		
				Set cashInputCreate = New_CashInput(1, 1, 0)
				With cashInputCreate
								.commonTab.office = "00"
        .commonTab.department = "1"
        .commonTab.date = "250122"
        .commonTab.dateForCheck = "25/01/22"
        .commonTab.cashRegister = "012"
        .commonTab.cashRegisterAcc = "000003701  "
        .commonTab.curr = "001"
        .commonTab.accCredit = "33170080101"
        .commonTab.amount = "46898.89"
        .commonTab.amountForCheck = "46,898.89"
        .commonTab.cashierChar = "023"
        .commonTab.base = "MC ëï³Ý¹³ñï-550****00****0*3;"
        .commonTab.aim = "Ð³Ù³Ó³ÛÝ Ñ³ßÇí-³åñ³Ýù³·ñÇ Ã."
        .commonTab.payer = "00000233"
        .commonTab.name = "§Ð³ÛÏ³Ï³Ý Ìñ³·ñ»ñ¦"
        .commonTab.surname = "êäÀ"
        .commonTab.id = "AM00000233"
        .commonTab.idForCheck = "AM00000233"
        .commonTab.idType = "06"
        .commonTab.idTypeForCheck = "06"
        .commonTab.idGivenBy = "021"
        .commonTab.idGivenByForCheck = "021"
        .commonTab.idGiveDate = "16022004"
        .commonTab.idGiveDateForCheck = "16/02/2004"
        .commonTab.idValidUntil = "16022024"
        .commonTab.idValidUntilForCheck = "16/02/2024"
        .commonTab.birthDate = "05081992"
        .commonTab.birthDateForCheck = "05/08/1992"
        .commonTab.citizenship = "2"
        .commonTab.country = "JP"
        .commonTab.residence = "010010239"
        .commonTab.city = "Ü³·áÛ³ "
        .commonTab.street = "öáÕáó 2"
        .commonTab.apartment = "´Ý³Ï³ñ³Ý 5"
        .commonTab.house = "Þ»Ýù 4/3"
        .commonTab.email = "japanesMen@mail.ru"
        .commonTab.emailForCheck = "japanesMen@mail.ru"
        .chargeTab.office = .commonTab.office
        .chargeTab.department = .commonTab.department
        .chargeTab.chargeAcc = "000003701  "
        .chargeTab.chargeAccForCheck = "77782963313"
        .chargeTab.chargeCurr = "001"
        .chargeTab.chargeCurrForCheck = "001"
        .chargeTab.cbExchangeRate = "400.0000/1"
        .chargeTab.chargeType = "02"
        .chargeTab.chargeAmount = "234.49"
        .chargeTab.chargeAmoForCheck = "234.49"
        .chargeTab.chargePercent = "0.5000"
        .chargeTab.chargePerForCheck = "0.5000"
        .chargeTab.incomeAcc = "000527500  "
        .chargeTab.incomeAccCurr = "000"
        .chargeTab.buyAndSell = "1"
        .chargeTab.buyAndSellForCheck = "1"
        .chargeTab.operType = "1"
        .chargeTab.operPlace = "3"
        .chargeTab.operArea = "9X"
        .chargeTab.operAreaForCheck = "9X"
        .chargeTab.nonResident = 0
        .chargeTab.nonResidentForCheck = 0
        .chargeTab.legalStatus = "11"
        .chargeTab.legalStatusForCheck = "11"
        .chargeTab.comment = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
        .chargeTab.commentForCheck = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
        .chargeTab.clientAgreeData = "îìÚ²ÈÜºñ"
        .coinTab.coin = "0.89"
        .coinTab.coinForCheck = "0.00"
        .coinTab.coinPayCurr = "000"
        .coinTab.coinBuyAndSell = "2"
        .coinTab.coinPayAcc = "000003700  "
        .coinTab.coinExchangeRate = "370.0000/1"
        .coinTab.coinCBExchangeRate = "400.0000/1"
        .coinTab.coinPayAmount = "329.30"
        .coinTab.coinPayAmountForCheck = "329.30"
        .coinTab.amountWithMainCurr = "46,898.00"
        .coinTab.amountCurrForCheck = "46,898.00"
        .coinTab.incomeOutChange = "000931900  "
        .coinTab.damagesOutChange = "001434300  "
        .attachedTab.linkName(0) = "attachedLink_1"
        .attachedTab.addFiles(0) = Project.Path & "Stores\Attach file\Photo.jpg"
        .attachedTab.fileName(0) = "Photo.jpg"
        .attachedTab.addLinks(0) = Project.Path & "Stores\Attach file\Photo.jpg"
				End With
				
    Set cashInputEdit = New_CashInput(1, 0, 1)
				With cashInputEdit
        .commonTab.office = "00"
        .commonTab.department = "1"
        .commonTab.date = "280122"
        .commonTab.dateForCheck = "28/01/22"
        .commonTab.cashRegister = "001"
        .commonTab.cashRegisterAcc = "000001101  "
        .commonTab.curr = "001"
        .commonTab.accCredit = "33170080101"
        .commonTab.amount = "8165.85"
        .commonTab.amountForCheck = "8,165.85"
        .commonTab.cashierChar = "04 "
        .commonTab.base = "MC ëï³Ý¹³ñï"
        .commonTab.aim = "Ð³Ù³Ó³ÛÝ Ñ³ßÇí"
        .commonTab.payer = "![End][Del]" 
        .commonTab.name = "Ð³ÛÏ³Ï³Ý Ìñ³·ñ»ñ êäÀ"
        .chargeTab.office = .commonTab.office
        .chargeTab.department = .commonTab.department
        .chargeTab.chargeAcc = "000001101  "
        .chargeTab.chargeAccForCheck = "000001101  "
        .chargeTab.chargeCurr = "001"
        .chargeTab.chargeCurrForCheck = "001"
        .chargeTab.cbExchangeRate = "400.0000/1"
        .chargeTab.chargeType = "05"
        .chargeTab.chargeAmount = "4.20"
        .chargeTab.chargeAmoForCheck = "4.20"
        .chargeTab.chargePercent = "0.0514"
        .chargeTab.chargePerForCheck = "0.0514"
        .chargeTab.incomeAcc = "000468200  "
        .chargeTab.incomeAccCurr = "000"
        .chargeTab.buyAndSell = "1"
        .chargeTab.buyAndSellForCheck = "1"
        .chargeTab.operType = "1"
        .chargeTab.operPlace = "3"
        .chargeTab.operArea = "9P"
        .chargeTab.operAreaForCheck = "9X"
        .chargeTab.nonResident = 0
        .chargeTab.nonResidentForCheck = 0
        .chargeTab.legalStatus = "12"
        .chargeTab.legalStatusForCheck = "11"
        .chargeTab.comment = "²ñï³ñÅ.·³ÝÓáõÙ"
        .chargeTab.commentForCheck = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
        .chargeTab.clientAgreeData = "îìÚ²ÈÜºñ 002"
        .coinTab.coin = "0.85"
        .coinTab.coinForCheck = "0.89"
        .coinTab.coinPayCurr = "000"
        .coinTab.coinBuyAndSell = "2"
        .coinTab.coinPayAcc = "000001100  "
        .coinTab.coinExchangeRate = "370.0000/1"
        .coinTab.coinCBExchangeRate = "400.0000/1"
        .coinTab.coinPayAmount = "314.50"
        .coinTab.coinPayAmountForCheck = "314.50"
        .coinTab.amountWithMainCurr = "8,165.00"
        .coinTab.amountCurrForCheck = "8,165.00"
        .coinTab.incomeOutChange = "000931900  "
        .coinTab.damagesOutChange = "001434300  "
        .attachedTab.addFiles(0) = Project.Path & "Stores\Attach file\Photo.jpg"
        .attachedTab.fileName(0) = "Photo.jpg"
        .attachedTab.delFiles(0) = Project.Path & "Stores\Attach file\Photo.jpg"
				End With
				
				Set workingDocs = New_MainAccWorkingDocuments()
				With workingDocs
								.startDate = cashInputCreate.commonTab.date
								.endDate = cashInputEdit.commonTab.date
				End With
				
				Set verifyDoc = New_VerificationDocument()
				verifyDoc.DocType = "KasPrOrd"
End Sub

Sub DB_Initialize()		
    Set dbo_FOLDERS(0) = New_DB_FOLDERS()
    With dbo_FOLDERS(0)
        .fKEY = cashInputCreate.fIsn
        .fISN = cashInputCreate.fIsn
        .fNAME = "KasPrOrd"
        .fSTATUS = "1"
        .fFOLDERID = ".D.GlavBux "
        .fCOM = "Î³ÝËÇÏ Ùáõïù"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
End	Sub

Sub Check_DB_Draft()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputCreate.fIsn, 2)
    Call CheckDB_DOCLOG(cashInputCreate.fIsn, "77", "N", "0", "", 1)
    Call CheckDB_DOCLOG(cashInputCreate.fIsn, "77", "F", "0", "", 1)
    
    'SQL Ստուգում DOCSATTACH աղուսյակում 
    Log.Message "SQL Ստուգում DOCSATTACH աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSATTACH", "fISN", cashInputCreate.fIsn, 2)
    Call CheckDB_DOCSATTACH(cashInputCreate.fIsn, Project.Path & "Stores\Attach file\Photo.jpg", "1", "attachedLink_1                                    ", 1)
    Call CheckDB_DOCSATTACH(cashInputCreate.fIsn, "Photo.jpg", "0", "", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    fBODY = " ACSBRANCH:00 ACSDEPART:1 TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 USERID:  77 DOCNUM:" & cashInputCreate.commonTab.docNum & " DATE:20220125 KASSA:012 ACCDB:000003701 CUR:001 ACCCR:33170080101 SUMMA:46898.89 KASSIMV:023 BASE:MC ëï³Ý¹³ñï-550****00****0*3; AIM:Ð³Ù³Ó³ÛÝ Ñ³ßÇí-³åñ³Ýù³·ñÇ Ã. CLICODE:00000233 PAYER:§Ð³ÛÏ³Ï³Ý Ìñ³·ñ»ñ¦ PAYERLASTNAME:êäÀ PASSNUM:AM00000233 PASTYPE:06 PASBY:021 DATEPASS:20040216 DATEEXPIRE:20240216 DATEBIRTH:19920805 CITIZENSHIP:2 COUNTRY:JP COMMUNITY:010010239 CITY:Ü³·áÛ³ APARTMENT:´Ý³Ï³ñ³Ý 5 ADDRESS:öáÕáó 2 BUILDNUM:Þ»Ýù 4/3 EMAIL:japanesMen@mail.ru ACSBRANCHINC:00 ACSDEPARTINC:1 CHRGACC:000003701 TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 CHRGCUR:001 CHRGCBCRS:400.0000/1 PAYSCALE:02 CHRGSUM:234.49 PRSNT:0.5 CHRGINC:000527500 CUPUSA:1 CURTES:1 CURVAIR:3 VOLORT:9X NONREZ:0 JURSTAT:11 COMM:²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ AGRDETAILS:îìÚ²ÈÜºñ XSUM:0.89 XCUR:000 XACC:000003700 XDLCRS:370/1 XDLCRSNAME:000 / 001 XCBCRS:400.0000/1 XCBCRSNAME:000 / 001 XCUPUSA:2 XCURSUM:329.3 XSUMMAIN:46898 XINC:000931900 XEXP:001434300 USEOVERLIMIT:0 NOTSENDABLE:0  "  
    fBODY = Replace(fBODY, " ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", cashInputCreate.fIsn, 1)
    Call CheckDB_DOCS(cashInputCreate.fIsn, "KasPrOrd", "0", fBODY, 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputCreate.fIsn, 1)
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
End	Sub

Sub Check_DB_Create()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputCreate.fIsn, 3)
    Call CheckDB_DOCLOG(cashInputCreate.fIsn, "77", "E", "2", "", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    fBODY = " ACSBRANCH:00 ACSDEPART:1 BLREP:0 OPERTYPE:MSC TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 USERID:  77 DOCNUM:" & cashInputCreate.commonTab.docNum & " DATE:20220125 KASSA:012 ACCDB:000003701 CUR:001 ACCCR:33170080101 SUMMA:46898.89 KASSIMV:023 BASE:MC ëï³Ý¹³ñï-550****00****0*3; AIM:Ð³Ù³Ó³ÛÝ Ñ³ßÇí-³åñ³Ýù³·ñÇ Ã. CLICODE:00000233 PAYER:§Ð³ÛÏ³Ï³Ý Ìñ³·ñ»ñ¦ PAYERLASTNAME:êäÀ PASSNUM:AM00000233 PASTYPE:06 PASBY:021 DATEPASS:20040216 DATEEXPIRE:20240216  DATEBIRTH:19920805 CITIZENSHIP:2 COUNTRY:JP COMMUNITY:010010239 CITY:Ü³·áÛ³ APARTMENT:´Ý³Ï³ñ³Ý 5 ADDRESS:öáÕáó 2 BUILDNUM:Þ»Ýù 4/3 EMAIL:japanesMen@mail.ru ACSBRANCHINC:00 ACSDEPARTINC:1 CHRGACC:000003701 TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 CHRGCUR:001 CHRGCBCRS:400.0000/1 PAYSCALE:02 CHRGSUM:234.49 PRSNT:0.5 CHRGINC:000527500 CUPUSA:1 CURTES:1 CURVAIR:3 VOLORT:9X NONREZ:0 JURSTAT:11 COMM:²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ AGRDETAILS:îìÚ²ÈÜºñ XSUM:0.89 XCUR:000 XACC:000003700 XDLCRS:   370.0000/    1 XDLCRSNAME:000 / 001 XCBCRS:400.0000/1 XCBCRSNAME:000 / 001 XCUPUSA:2 XCURSUM:329.3 XSUMMAIN:46898 XINC:000931900 XEXP:001434300 USEOVERLIMIT:0 NOTSENDABLE:0  "   
    fBODY = Replace(fBODY, " ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", cashInputCreate.fIsn, 1)
    Call CheckDB_DOCS(cashInputCreate.fIsn, "KasPrOrd", "2", fBODY, 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputCreate.fIsn, 2)
    With dbo_FOLDERS(0)
        .fSTATUS = "5"
        .fFOLDERID = "C.1732628"
        .fSPEC = "²Ùë³ÃÇí- 25/01/22 N- " & cashInputCreate.commonTab.docNum & " ¶áõÙ³ñ-            46,898.89 ²ñÅ.- 001 [Üáñ]"
        .fECOM = "Cash Deposit Advice"
        .fDCBRANCH = ""
        .fDCDEPART = ""
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
    Set dbo_FOLDERS(1) = New_DB_FOLDERS()
    With dbo_FOLDERS(1)
        .fKEY = cashInputCreate.fIsn
        .fISN = cashInputCreate.fIsn
        .fNAME = "KasPrOrd"
        .fSTATUS = "5"
        .fFOLDERID = "Oper.20220125"
        .fCOM = "Î³ÝËÇÏ Ùáõïù"
        .fSPEC = cashInputCreate.commonTab.docNum & "77700000003701  7770033170080101        46898.89001Üáñ                                                   77§Ð³ÛÏ³Ï³Ý Ìñ³·ñ»ñ¦ êäÀ          AM00000233 021 16/02/2004                                       Ð³Ù³Ó³ÛÝ Ñ³ßÇí-³åñ³Ýù³·ñÇ Ã. MC ëï³Ý¹³ñï-550****00****0*3;                                                                                  "
        .fECOM = "Cash Deposit Advice"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERS(1), 1)
  
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", cashInputCreate.fIsn, 8)
    Call Check_HI_CE_accounting ("2022-01-25", cashInputCreate.fIsn, "11", "1630475", "18759200.00", "001", "46898.00", "MSC", "D")
    Call Check_HI_CE_accounting ("2022-01-25", cashInputCreate.fIsn, "11", "707641461", "18759200.00", "001", "46898.00", "MSC", "C")
    Call Check_HI_CE_accounting ("2022-01-25", cashInputCreate.fIsn, "11", "1629708", "26.70", "000", "26.70", "MSC", "D")
    Call Check_HI_CE_accounting ("2022-01-25", cashInputCreate.fIsn, "11", "707641461", "26.70", "001", "0.00", "MSC", "C")
    Call Check_HI_CE_accounting ("2022-01-25", cashInputCreate.fIsn, "11", "1630474", "329.30", "000", "329.30", "CEX", "D")
    Call Check_HI_CE_accounting ("2022-01-25", cashInputCreate.fIsn, "11", "707641461", "329.30", "001", "0.89", "CEX", "C")
    Call Check_HI_CE_accounting ("2022-01-25", cashInputCreate.fIsn, "11", "1629291", "93796.00", "000", "93796.00", "FEX", "C")
    Call Check_HI_CE_accounting ("2022-01-25", cashInputCreate.fIsn, "11", "1630475", "93796.00", "001", "234.49", "FEX", "D")
End Sub

Sub Check_DB_Edit()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputEdit.fIsn, 4)
    Call CheckDB_DOCLOG(cashInputEdit.fIsn, "77", "E", "2", "", 2)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    fBODY = " ACSBRANCH:00 ACSDEPART:1 BLREP:0 OPERTYPE:MSC TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 USERID:  77 DOCNUM:" & cashInputEdit.commonTab.docNum & " DATE:20220128 KASSA:001 ACCDB:000001101 CUR:001 ACCCR:33170080101 SUMMA:8165.85 KASSIMV:04 BASE:MC ëï³Ý¹³ñï AIM:Ð³Ù³Ó³ÛÝ Ñ³ßÇí PAYER:Ð³ÛÏ³Ï³Ý Ìñ³·ñ»ñ êäÀ ACSBRANCHINC:00 ACSDEPARTINC:1 CHRGACC:000001101 TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 CHRGCUR:001 CHRGCBCRS:400.0000/1 PAYSCALE:05 CHRGSUM:4.2 PRSNT:0.0514 CHRGINC:000468200 CUPUSA:1 CURTES:1 CURVAIR:3 VOLORT:9P NONREZ:0 JURSTAT:12 COMM:²ñï³ñÅ.·³ÝÓáõÙ AGRDETAILS:îìÚ²ÈÜºñ 002 XSUM:0.85 XCUR:000 XACC:000001100 XDLCRS:370/1 XDLCRSNAME:000 / 001 XCBCRS:400.0000/1 XCBCRSNAME:000 / 001 XCUPUSA:2 XCURSUM:314.5 XSUMMAIN:8165 XINC:000931900 XEXP:001434300 USEOVERLIMIT:0 NOTSENDABLE:0  "  
    fBODY = Replace(fBODY, " ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", cashInputEdit.fIsn, 1)
    Call CheckDB_DOCS(cashInputEdit.fIsn, "KasPrOrd", "2", fBODY, 1)
  
    'SQL Ստուգում DOCSATTACH աղուսյակում 
    Log.Message "SQL Ստուգում DOCSATTACH աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSATTACH", "fISN", cashInputEdit.fIsn, 1)
    Call CheckDB_DOCSATTACH(cashInputEdit.fIsn, "Photo.jpg", "0", "", 1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputEdit.fIsn, 1)
    With dbo_FOLDERS(1)
        .fFOLDERID = "Oper.20220128"
        .fSPEC = cashInputEdit.commonTab.docNum & "77700000001101  7770033170080101         8165.85001ÊÙµ³·ñíáÕ                                             77Ð³ÛÏ³Ï³Ý Ìñ³·ñ»ñ êäÀ                                                                            Ð³Ù³Ó³ÛÝ Ñ³ßÇí MC ëï³Ý¹³ñï                                                                                                                  "
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERS(1), 1)
  
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", cashInputEdit.fIsn, 8)
    Call Check_HI_CE_accounting ("2022-01-28", cashInputEdit.fIsn, "11", "1630171", "3266000.00", "001", "8165.00", "MSC", "D")
    Call Check_HI_CE_accounting ("2022-01-28", cashInputEdit.fIsn, "11", "707641461", "3266000.00", "001", "8165.00", "MSC", "C")
    Call Check_HI_CE_accounting ("2022-01-28", cashInputEdit.fIsn, "11", "1629708", "25.50", "000", "25.50", "MSC", "D")
    Call Check_HI_CE_accounting ("2022-01-28", cashInputEdit.fIsn, "11", "707641461", "25.50", "001", "0.00", "MSC", "C")
    Call Check_HI_CE_accounting ("2022-01-28", cashInputEdit.fIsn, "11", "1630170", "314.50", "000", "314.50", "CEX", "D")
    Call Check_HI_CE_accounting ("2022-01-28", cashInputEdit.fIsn, "11", "707641461", "314.50", "001", "0.85", "CEX", "C")
    Call Check_HI_CE_accounting ("2022-01-28", cashInputEdit.fIsn, "11", "1629232", "1680.00", "000", "1680.00", "FEX", "C")
    Call Check_HI_CE_accounting ("2022-01-28", cashInputEdit.fIsn, "11", "1630171", "1680.00", "001", "4.20", "FEX", "D")
End Sub

Sub Check_DB_SendToVerify()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputEdit.fIsn, 5)
    Call CheckDB_DOCLOG(cashInputEdit.fIsn, "77", "M", "101", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", cashInputEdit.fIsn, 1)
    Call CheckDB_DOCS(cashInputEdit.fIsn, "KasPrOrd", "101", fBODY, 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputEdit.fIsn, 2)
    With dbo_FOLDERS(0)
        .fFOLDERID = "Oper.20220128"
        .fSTATUS = "0"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
        .fSPEC = cashInputEdit.commonTab.docNum & "77700000001101  7770033170080101         8165.85001àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý                                 77Ð³ÛÏ³Ï³Ý Ìñ³·ñ»ñ êäÀ                                            001                             Ð³Ù³Ó³ÛÝ Ñ³ßÇí MC ëï³Ý¹³ñï                                                                                                                  "
    End With 
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
    With dbo_FOLDERS(1)
        .fKEY = cashInputEdit.fIsn
        .fISN = cashInputEdit.fIsn
        .fNAME = "KasPrOrd"
        .fSTATUS = "4"
        .fFOLDERID = "Ver.20220128001"
        .fCOM = "Î³ÝËÇÏ Ùáõïù"
        .fSPEC = cashInputEdit.commonTab.docNum & "77700000001101  7770033170080101         8165.85001  77Ð³Ù³Ó³ÛÝ Ñ³ßÇí                  MC ëï³Ý¹³ñï                     Ð³ÛÏ³Ï³Ý Ìñ³·ñ»ñ êäÀ            "
        .fECOM = "Cash Deposit Advice"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With 
    Call CheckDB_FOLDERS(dbo_FOLDERS(1), 1)
End Sub

Sub Check_DB_Validate()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputEdit.fIsn, 7)
    Call CheckDB_DOCLOG(cashInputEdit.fIsn, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(cashInputEdit.fIsn, "77", "C", "15", "", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", cashInputEdit.fIsn, 1)
    Call CheckDB_DOCS(cashInputEdit.fIsn, "KasPrOrd", "15", fBODY, 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputEdit.fIsn, 1)
    With dbo_FOLDERS(0)
        .fSTATUS = "4"
        .fSPEC = cashInputEdit.commonTab.docNum & "77700000001101  7770033170080101         8165.85001Ð³ëï³ïí³Í                                             77Ð³ÛÏ³Ï³Ý Ìñ³·ñ»ñ êäÀ                                                                            Ð³Ù³Ó³ÛÝ Ñ³ßÇí MC ëï³Ý¹³ñï                                                                                                                  "
    End With 
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
End Sub

Sub Check_DB_Delete()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputEdit.fIsn, 8)
    Call CheckDB_DOCLOG(cashInputEdit.fIsn, "77", "D", "999", "", 1)
				
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", cashInputEdit.fIsn, 1)
    Call CheckDB_DOCS(cashInputEdit.fIsn, "KasPrOrd", "999", fBODY, 1)
		
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    With dbo_FOLDERS(0)
        .fNAME = "KasPrOrd"
        .fSTATUS = "0"
        .fFOLDERID = ".R." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d")
        .fSPEC = Left_Align(Get_Compname_DOCLOG(cashInputEdit.fIsn), 16) & "GlavBux ARMSOFT                       0115 "
        .fCOM = ""
        .fECOM = ""
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With 
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputEdit.fIsn, 1)
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
End Sub