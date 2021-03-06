'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT Library_CheckDB
'USEUNIT OLAP_Library
'USEUNIT Card_Library
'USEUNIT Mortgage_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT DAHK_Library_Filter
'USEUNIT Library_Contracts
'USEUNIT Overlimit_Library
'USEUNIT CashOutput_Confirmpases_Library
'USEUNIT Main_Accountant_Filter_Library
'USEUNIT CashInput_Confirmphases_Library

'USEUNIT Constants
Option Explicit

'Test Case Id - 169213

Dim dbFOLDERS(8),dbHI2(2)
    
Sub Cash_Accounting_ByCashRequest()
  
    Dim sDATE,eDATE
    Dim CashAccountingObject,CashInIsn,VerificationDoc
    Dim CashRequestIsn,CashRequestObject,NewCashRequests
    Dim CashOutObject,CashOutIsn,CashIn
    
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    sDATE = "20030101"
    eDATE = "20260101"
    Call Initialize_AsBank("bank", sDATE, eDATE)
    Login("ARMSOFT")
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''-- Փոխել CASHACCFLAG Պարամետրի արժեքը --''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Փոխել CASHACCFLAG Պարամետրի արժեքը --" ,,, DivideColor     
    
    'Մուտք Ադմինիստրատորի ԱՇՏ 4.0
    Call ChangeWorkspace(c_Admin40)
    
    Call GoTo_SystemParameters("CASHACCFLAG") 
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToEdit)
    wMDIClient.Refresh
    
    'Լրացնել "Արժեք" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "CheckBox", "VALUE", "1")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    BuiltIn.Delay(2000)
    Call Close_Pttel("frmPttel")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''-- Հաշիվներ թղթապանակից կատարել "Կանխիկ մուտք" գործողությունը --'''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հաշիվներ թղթապանակից կատարել Կանխիկ մուտք գործողությունը --" ,,, DivideColor 
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ
    Call ChangeWorkspace(c_ChiefAcc)
    Call OpenAccauntsFolder("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³ßÇíÝ»ñ","1","","00057133311","","","","","","",1,"","","","","",1,1,1,"","","","","","ACCS","0")
    Call CheckPttel_RowCount("frmPttel", 1)            
                                                  
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_InnerOpers &"|"& c_Cashin)             
      
    Set CashIn = New_CashIn   
    With CashIn
        .Date = "030320"
        .Amount = "550005.06"
        .CashLabel = "051"
        .Base = "Ð³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ"
        .Aim = "Ð³Ù³Ó³ÛÝ Ã. Ñ³ßíÇ"
        .Depositor = "00000075"
        .FirstName = "master"
        .LastName = "1_²ÝáõÝ_²ÝáõÝ_²ÝáõÝ_²ÝáõÝ_²ÝáõÝ_²ÝáõÝ_²ÝáõÝ_²ÝáõÝ_²ÝáõÝ_²ÝáõÝ_²ÝáõÝ_22"
        .IdNumber = ""
        .Issued = "005"
        .IssuedDate = "010103"
        .DateOfExpire = "010126"
        .DateOfBirth = "260797"
        .Citizenship = "1"
        .Country = "AM"
        .Community = "010010130"
        .City = "Yerevan"
        .Flat = "1212112111"
        .Street = "Mikoyan_Mikoyan_Mikoyan_Mikoyan_Mikoyan_Mikoyan_Mikoyan_Mikoyan_Mikoyan1"
        .House = "1222222222"
        .Email = "1aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa@mail.ru"
        .NonResident = 0
        .Comment = ""
        .CliAgrDetails = "Clients_Agrients_Agre_1"
        .SubAmount = "111.01"
        .SubAmountToBePaid = "41,073.70"
        .AmountInPrimaryCurr = "549,894.05"
    End With 

    CashInIsn = Fill_CashIn(CashIn)
    
    'Եթե քաղվածքի պատուհանը հայտնվել է, ապա փակում է
    If wMDIClient.VBObject("FrmSpr").Exists Then
        wMDIClient.VBObject("FrmSpr").Close
    Else
        Log.Error "Statement window doesn't exist!",,,ErrorColor
    End If
    Call Close_Pttel("frmPttel")

    Log.Message "fISN = " & CashInIsn ,,,SqlDivideColor
    
    Call SQL_Initialize_Cash_Acc_ByCashReq(CashInIsn)
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  ACSBRANCH:00  ACSDEPART:1  BLREP:0  OPERTYPE:MSC  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  USERID:  77  DOCNUM:"&CashIn.DocNum&"  DATE:20200303  KASSA:001  ACCDB:000001101  CUR:001  ACCCR:00057133311  SUMMA:550005.06  KASSIMV:021  BASE:Ð³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ  AIM:Ð³Ù³Ó³ÛÝ Ã. Ñ³ßíÇ  CLICODE:00000075  PAYER:àõÇÉÉÇ³Ù  PAYERLASTNAME:¶»ÛÃë  PASSNUM:N1234567890  PASBY:005  DATEPASS:20030101  DATEEXPIRE:20260101  CITIZENSHIP:1  COUNTRY:AM  ADDRESS:ºñ¨³Ý, ÐÐ, ²ëï³ýÛ³Ý 1,1  ACSBRANCHINC:00  ACSDEPARTINC:1  CHRGACC:000001100  TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  CHRGCUR:000  CHRGCBCRS:1/1  CURTES:1  CURVAIR:3  VOLORT:7  NONREZ:0  JURSTAT:21  XSUM:111.01  XCUR:000  XACC:000001100  XDLCRS:   370.0000/    1  XDLCRSNAME:000 / 001  XCBCRS:400.0000/1  XCBCRSNAME:000 / 001  XCUPUSA:2  XCURSUM:41073.7  XSUMMAIN:549894.05  XINC:000931900  XEXP:001434300  USEOVERLIMIT:0  NOTSENDABLE:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",CashInIsn,1)
    Call CheckDB_DOCS(CashInIsn,"KasPrOrd","2",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",CashInIsn,2)
    Call CheckDB_DOCLOG(CashInIsn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(CashInIsn,"77","C","2"," ",1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    dbFOLDERS(1).fSPEC = "²Ùë³ÃÇí- 03/03/20 N- "&CashIn.DocNum&" ¶áõÙ³ñ-           550,005.06 ²ñÅ.- 001 [Üáñ]"
    dbFOLDERS(2).fSPEC = CashIn.DocNum&"77700000001101  7770000057133311       550005.06001Üáñ                                                   77àõÇÉÉÇ³Ù ¶»ÛÃë                  N1234567890 005 01/01/2003                                      Ð³Ù³Ó³ÛÝ Ã. Ñ³ßíÇ Ð³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ                                                                                                       "
    
    Call CheckQueryRowCount("FOLDERS","fISN",CashInIsn,2)
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)
    
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",CashInIsn,6)                       
    Call Check_HI_CE_accounting ("20200303",CashInIsn, "11", "1630171","219957620.00", "001", "549894.05", "MSC", "D")
    Call Check_HI_CE_accounting ("20200303",CashInIsn, "11", "1714909","219957620.00", "001", "549894.05", "MSC", "C")
    Call Check_HI_CE_accounting ("20200303",CashInIsn, "11", "1629708","3330.30", "000", "3330.30", "MSC", "D")
    Call Check_HI_CE_accounting ("20200303",CashInIsn, "11", "1714909","3330.30", "001", "0.00", "MSC", "C")
    Call Check_HI_CE_accounting ("20200303",CashInIsn, "11", "1630170","41073.70", "000", "41073.70", "CEX", "D")
    Call Check_HI_CE_accounting ("20200303",CashInIsn, "11", "1714909","41073.70", "001", "111.01", "CEX", "C")

    'SQL Ստուգում HI2 աղուսյակում 
    Call CheckQueryRowCount("HI2","fBASE",CashInIsn,1)
    Call CheckDB_HI2(dbHI2(1),1)

    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckDB_HIREST("11", "1630171","1282777522.50","001","2819456.41",1)
    Call CheckDB_HIREST("11", "1714909","-220002024.00","001","-550005.06",1)    
    Call CheckDB_HIREST("11", "1629708","11811.40","000","11811.40",1)   
    Call CheckDB_HIREST("11", "1630170","83505716.30","000","83505716.30",1)   
    
    'SQL Ստուգում HIREST2 աղուսյակում
    Call CheckDB_HIREST2("10","1628383","1578250","0.00","001","550105.06", 1)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''-- "Աշխատանքային փաստաթղթեր" թղթապանակից կատարել "Ուղարկել հաստատման" գործողությունը --'''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կատարել Ուղարկել հաստատման գործողությունը --",,,DivideColor       
    
    wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
    'Լրացնել "Ամսաթիվ" դաշտերը
    Call Rekvizit_Fill("Dialog",1,"General","PERN", "030320")
    Call Rekvizit_Fill("Dialog",1,"General","PERK", "030320")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    If WaitForPttel("frmPttel") Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_SendToVer)
        BuiltIn.Delay(2000)
        Call MessageExists(2,"àõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý")
        Call ClickCmdButton(2, "Î³ï³ñ»É")
        Call Close_Pttel("frmPttel") 
    Else
        Log.Error "Can Not Open Աշխատանքային փաստաթղթեր pttel",,,ErrorColor      
    End If  
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  ACSBRANCH:00  ACSDEPART:1  BLREP:0  OPERTYPE:MSC  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  USERID:  77  DOCNUM:"&CashIn.DocNum&"  DATE:20200303  KASSA:001  ACCDB:000001101  CUR:001  ACCCR:00057133311  SUMMA:550005.06  KASSIMV:021  BASE:Ð³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ  AIM:Ð³Ù³Ó³ÛÝ Ã. Ñ³ßíÇ  CLICODE:00000075  PAYER:àõÇÉÉÇ³Ù  PAYERLASTNAME:¶»ÛÃë  PASSNUM:N1234567890  PASBY:005  DATEPASS:20030101  DATEEXPIRE:20260101  CITIZENSHIP:1  COUNTRY:AM  ADDRESS:ºñ¨³Ý, ÐÐ, ²ëï³ýÛ³Ý 1,1  ISTLLCREATED:1  ACSBRANCHINC:00  ACSDEPARTINC:1  CHRGACC:000001100  TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  CHRGCUR:000  CHRGCBCRS:1/1  CURTES:1  CURVAIR:3  VOLORT:7  NONREZ:0  JURSTAT:21  XSUM:111.01  XCUR:000  XACC:000001100  XDLCRS:   370.0000/    1  XDLCRSNAME:000 / 001  XCBCRS:400.0000/1  XCBCRSNAME:000 / 001  XCUPUSA:2  XCURSUM:41073.7  XSUMMAIN:549894.05  XINC:000931900  XEXP:001434300  USEOVERLIMIT:0  NOTSENDABLE:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",CashInIsn,1)
    Call CheckDB_DOCS(CashInIsn,"KasPrOrd","101",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",CashInIsn,3)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    dbFOLDERS(1).fSTATUS = "0"
    dbFOLDERS(1).fSPEC = "²Ùë³ÃÇí- 03/03/20 N- "&CashIn.DocNum&" ¶áõÙ³ñ-           550,005.06 ²ñÅ.- 001 [àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý]"
    dbFOLDERS(2).fSTATUS = "0"
    dbFOLDERS(2).fSPEC = CashIn.DocNum&"77700000001101  7770000057133311       550005.06001àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý                                 77àõÇÉÉÇ³Ù ¶»ÛÃë                                                  001                             Ð³Ù³Ó³ÛÝ Ã. Ñ³ßíÇ Ð³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ                                                                                                       "
    dbFOLDERS(3).fSPEC = CashIn.DocNum&"77700000001101  7770000057133311       550005.06001  77Ð³Ù³Ó³ÛÝ Ã. Ñ³ßíÇ               Ð³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ             àõÇÉÉÇ³Ù ¶»ÛÃë"
    
    Call CheckQueryRowCount("FOLDERS","fISN",CashInIsn,3)
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)
    Call CheckDB_FOLDERS(dbFOLDERS(3),1)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''-- Գլխավոր հաշվապահ/Հաստատվող փաստաթղթեր(|) թղթապանակից կատարել "Վավերացնել" գործողությունը --''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կատարել Վավերացնել գործողությունը --",,,DivideColor   

    Set VerificationDoc = New_VerificationDocument()
        VerificationDoc.DocType = "KasPrOrd"
    
    Call GoToVerificationDocument("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ (I)",VerificationDoc) 
    
    If WaitForPttel("frmPttel") Then
        If SearchInPttel("frmPttel",7, "550005.06") Then
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(c_ToConfirm)
            BuiltIn.Delay(2000)
            Call ClickCmdButton(1, "Ð³ëï³ï»É")
        Else 
            Log.Error "Տողը չի գտնվել Հաստատվող փաստաթղթեր(|) թղթապանակում" ,,,ErrorColor
        End If
        Call Close_Pttel("frmPttel")
     Else
        Log.Error "Can Not Open Հաստատվող փաստաթղթեր(|) Window",,,ErrorColor      
     End If   
     
     'SQL Ստուգում DOCS աղուսյակում
    Call CheckQueryRowCount("DOCS","fISN",CashInIsn,1)
    Call CheckDB_DOCS(CashInIsn,"KasPrOrd","15",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",CashInIsn,5)
    Call CheckDB_DOCLOG(CashInIsn,"77","W","102"," ",1)
    Call CheckDB_DOCLOG(CashInIsn,"77","C","15"," ",1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    dbFOLDERS(1).fSTATUS = "4"
    dbFOLDERS(1).fSPEC = "²Ùë³ÃÇí- 03/03/20 N- "&CashIn.DocNum&" ¶áõÙ³ñ-           550,005.06 ²ñÅ.- 001 [Ð³ëï³ïí³Í]"
    dbFOLDERS(2).fSTATUS = "4"
    dbFOLDERS(2).fSPEC = CashIn.DocNum & "77700000001101  7770000057133311       550005.06001Ð³ëï³ïí³Í                                             77àõÇÉÉÇ³Ù ¶»ÛÃë                  N1234567890 005 01/01/2003                                      Ð³Ù³Ó³ÛÝ Ã. Ñ³ßíÇ Ð³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ                                                                                                       "
    
    Call CheckQueryRowCount("FOLDERS","fISN",CashInIsn,2)
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''-- "Աշխատանքային փաստաթղթեր" թղթապանակից կատարել "Վավերացնել" գործողությունը --'''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Աշխատանքային փաստաթղթեր թղթապանակից կատարել Վավերացնել գործողությունը --",,,DivideColor       
    
    wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
    'Լրացնել "Ամսաթիվ" դաշտերը
    Call Rekvizit_Fill("Dialog",1,"General","PERN", "030320")
    Call Rekvizit_Fill("Dialog",1,"General","PERK", "030320")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    If WaitForPttel("frmPttel") Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_ToConfirm)
        BuiltIn.Delay(2000)
        Call ClickCmdButton(1, "Ð³ëï³ï»É")
        Call Close_Pttel("frmPttel") 
    Else
        Log.Error "Can Not Open Աշխատանքային փաստաթղթեր pttel",,,ErrorColor      
    End If      

    'SQL Ստուգում DOCS աղուսյակում
    Call CheckQueryRowCount("DOCS","fISN",CashInIsn,1)
    Call CheckDB_DOCS(CashInIsn,"KasPrOrd","11",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",CashInIsn,7)
    Call CheckDB_DOCLOG(CashInIsn,"77","W","16"," ",1)
    Call CheckDB_DOCLOG(CashInIsn,"77","M","11","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Call CheckQueryRowCount("FOLDERS","fISN",CashInIsn,0)
    
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",CashInIsn,7)                       
    Call Check_HI_CE_accounting ("20200303",CashInIsn, "CE", "1559631","111.01", "000", "41073.70", "SAL", "D")
     
    'SQL Ստուգում PAYMENTS աղուսյակում  
    Call CheckQueryRowCount("PAYMENTS","fISN",CashInIsn,1)
    
    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckDB_HIREST("01", "1714909","-249928691.40","001","-621259.03",1)
    Call CheckDB_HIREST("01", "1629708","4060041.30","000","4060041.30",1)    
    Call CheckDB_HIREST("01", "1630170","1099353726.60","000","1099353726.60",1)   
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''-- Հաշիվներ թղթապանակից կատարել "Կանխիկացման հայտ" գործողությունը --'''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հաշիվներ թղթապանակից կատարել Կանխիկացման հայտ գործողությունը --" ,,, DivideColor 
    
    Call OpenAccauntsFolder("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³ßÇíÝ»ñ","1","","","","","","","00000075","",1,"","","","","",1,1,1,"","","","","","ACCS","0")
    Call CheckPttel_RowCount("frmPttel", 1)            
                                                  
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_InnerOpers &"|"& c_CashReq)            
    
    Set CashRequestObject = New_CashRequest()   
    With CashRequestObject
      .System = "1"
      .CashDivision = "00"
      .Department = "1"
      .RequestNumber = ""
      .CallDate = "040420"
      .CashDate = "050524"
      .Curr = "001"
      .Receiver = "00000075"
      .Amount = "1005600.01"
      .Comment = "CashRequest_CashRequest_CashRequest_1"
      .RecPaySystem = "Ð"
    End With  

    CashRequestIsn = Fill_CashRequest(CashRequestObject) 
    
    Log.Message "fISN = " & CashRequestIsn ,,,SqlDivideColor
    
    Call SQL_Initialize_Cash_Acc_ByCashReq(CashRequestIsn)
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  REMSYSTEM:1  ACSBRANCH:00  ACSDEPART:1  USERID:  77  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  DOCNUM:"&CashRequestObject.RequestNumber&"  DATE:20200404  REQUESTDATE:20240505  ACC:7770000057133311  ACCDB:00057133311  CURR:001  CLIENT:00000075  RECCLIENT:00000075  RECEIVER:àõÇÉÉÇ³Ù  RECEIVERLASTNAME:¶»ÛÃë  PASSNUM:N1234567890  PASTYPE:01  SUMMA:1005600.01  DESCR:CashRequest_CashRequest_CashRequest_1  CANCELREQ:0  PAYSYSIN:Ð  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",CashRequestIsn,1)
    Call CheckDB_DOCS(CashRequestIsn,"CBCshReq","2",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում
    Call CheckQueryRowCount("DOCLOG","fISN",CashRequestIsn,2)
    Call CheckDB_DOCLOG(CashRequestIsn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(CashRequestIsn,"77","C","2"," ",1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    
    dbFOLDERS(4).fSPEC = "²Ùë³ÃÇí- 05/05/24 N- "&CashRequestObject.RequestNumber&" ¶áõÙ³ñ-         1,005,600.01 ²ñÅ.- 001 [Üáñ]"
    dbFOLDERS(6).fSPEC = CashRequestObject.RequestNumber&"7770000057133311                      1005600.01001Üáñ                                                   77àõÇÉÉÇ³Ù ¶»ÛÃë                  N1234567890                                            Ð        CashRequest_CashRequest_CashRequest_1                                                                                                       "
    
    Call CheckQueryRowCount("FOLDERS","fISN",CashRequestIsn,3)
    Call CheckDB_FOLDERS(dbFOLDERS(4),1)
    Call CheckDB_FOLDERS(dbFOLDERS(5),1)
    Call CheckDB_FOLDERS(dbFOLDERS(6),1)
    
    Call Close_Pttel("frmPttel")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''-- Հաճախորդի սպասարկում և դրամարկղ ԱՇՏ/թղթապանակներ/Նոր ստեղծված կանխիկացման հայտ թղթապանակից կատարել հաստատել և ստեղծել կանխիկ ելք գործողությունները --''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Նոր ստեղծված կանխիկացման հայտ թղթապանակից կատարել հաստատել գործողությունը --" ,,, DivideColor 
    
    'Մուտք Հաճախորդի սպասարկում և դրամարկղ (ընդլայնված) ԱՇՏ
    Call ChangeWorkspace(c_CustomerService)
    
    Set NewCashRequests = New_NewCashRequest()
    With NewCashRequests
        .StartDate = "^A[Del]" & "010122"
        .EndDate = "^A[Del]" & "010125"
        .Division = "00"
        .Department = "1"
        .ClientsCode = "00000075"
        .Account = ""
        .Curr = "001"
     End With   
    
    Call GoTo_CustomerService_CashRequest(NewCashRequests)
        
    If WaitForPttel("frmPttel") Then
        If SearchInPttel("frmPttel",8, "1005600.01") Then
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(c_ToVerify)
            BuiltIn.Delay(2000)
            Call MessageExists(2,"Ð³ëï³ï»É")
            Call ClickCmdButton(5, "²Ûá")
            BuiltIn.Delay(2000)
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click("Ստեղծել «Կանխիկ Ելք»")
        Else 
            Log.Error "Տողը չի գտնվել Նոր ստեղծված կանխիկացման հայտ թղթապանակում" ,,,ErrorColor
        End If
     Else
        Log.Error "Can Not Open Նոր ստեղծված կանխիկացման հայտ Window",,,ErrorColor      
     End If 
    
    Set CashOutObject = New_CashOut   
    With CashOutObject
        .DocNum = ""
        .Date = "020221"
        .Amount = "1,000,500.08"
        .CashLabel = "051"
        .Depositor = "00000075"
        .FirstName = "àõÇÉÉÇ³Ù"
        .LastName = "¶»ÛÃë"
        .IdNumber = ""
        .Issued = "005"
        .IssuedDate = "010103"
        .DateOfExpire = "010126"
        .DateOfBirth = "260797"
        .Citizenship = "1"
        .Country = "AM"
        .Community = "010010130"
        .ChargesAccount = "000001100  "
        .Curr = "000"
        .ChargeType = "01"
        .ChargesAmount = "180,158.00"
        .Interest = "0.1000"
        .IncomeAccount = "000919400"
        .NonChargeableAmount = "550,105.06"
        .CliAgrDetails = "Clients_Agreement_Clients_Agre_1"
        .SubAmount = "22,222.22"
        .SubAmountToBePaid = "7,555,554.80"
        .AmountInPrimaryCurr = "978,277.86"
    End With 
    CashOutIsn = Fill_CashOut(CashOutObject)
    
    'Եթե քաղվածքի պատուհանը հայտնվել է, ապա փակում է
    If wMDIClient.VBObject("FrmSpr").Exists Then
        wMDIClient.VBObject("FrmSpr").Close
    Else
        Log.Error "Statement window doesn't exist!",,,ErrorColor
    End If
    Call Close_Pttel("frmPttel")
    
    Log.Message "DocNum = " & CashOutObject.DocNum,,,DivideColor2    
    Log.Message "fISN = " & CashOutIsn ,,,SqlDivideColor
     
    'SQL Ստուգում FOLDERS աղուսյակում 
    dbFOLDERS(5).fSPEC = "0000007500057133311      1005600.01001CashRequest_CashRequest_CashRequ0050Øß³Ïí³Í               77àõÇÉÉÇ³Ù ¶»ÛÃë                  William H. Gates                àõÇÉÉÇ³Ù ¶»ÛÃë                  "   
   
    Call CheckQueryRowCount("FOLDERS","fISN",CashRequestIsn,1)
    Call CheckDB_FOLDERS(dbFOLDERS(5),1)
    
    Call SQL_Initialize_Cash_Acc_ByCashReq(CashOutIsn)
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  ACSBRANCH:00  ACSDEPART:1  BLREP:0  OPERTYPE:MSC  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  USERID:  77  DOCNUM:"&CashOutObject.DocNum&"  DATE:20210202  ACCDB:00057133311  CUR:001  KASSA:001  ACCCR:000001101  SUMMA:1000500.08  TOTAL:1005600.01  KASSIMV:051  BASE:Î³ÝËÇÏ³óÙ³Ý Ñ³Ûï  AIM:CashRequest_CashRequest_CashRequest_1  CLICODE:00000075  RECEIVER:àõÇÉÉÇ³Ù  RECEIVERLASTNAME:¶»ÛÃë  PASSNUM:N1234567890  CITIZENSHIP:1  COUNTRY:AM  ADDRESS:ºñ¨³Ý, ÐÐ, ²ëï³ýÛ³Ý 1,1  FROMPAYORD:0  ACSBRANCHINC:00  ACSDEPARTINC:1  CHRGACC:000001100  TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  CHRGCUR:000  CHRGCBCRS:1/1  PAYSCALE:01  CHRGSUM:180158  PRSNT:0.1  CHRGINC:000919400  NOCRGSUM:550105.06  FRSHNOCRG:0  CURTES:1  CURVAIR:3  VOLORT:7  NONREZ:0  JURSTAT:21  COMM:¶³ÝÓáõÙ Ï³ÝËÇÏ³óáõÙÇó  PAYSYSIN:Ð  XSUM:22222.22  XCUR:000  XACC:000001100  XDLCRS:   340.0000/    1  XDLCRSNAME:000 / 001  XCBCRS:400.0000/1  XCBCRSNAME:000 / 001  XCUPUSA:1  XCURSUM:7555554.8  XSUMMAIN:978277.86  XINC:000931900  XEXP:001434300  SYSCASE:CashReq  NOTSENDABLE:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",CashOutIsn,1)
    Call CheckDB_DOCS(CashOutIsn,"KasRsOrd","2",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",CashOutIsn,2)
    Call CheckDB_DOCLOG(CashOutIsn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(CashOutIsn,"77","C","2"," ",1)
    
    'SQL Ստուգում DOCP աղուսյակում  
    Call CheckQueryRowCount("DOCP","fISN",CashOutIsn,1)
    Call CheckDB_DOCP(CashOutIsn,"KasRsOrd",CashRequestIsn,1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    dbFOLDERS(7).fSPEC = "²Ùë³ÃÇí- 02/02/21 N- "&CashOutObject.DocNum&" ¶áõÙ³ñ-         1,000,500.08 ²ñÅ.- 001 [Üáñ]"
    dbFOLDERS(8).fSPEC = CashOutObject.DocNum & "777000005713331177700000001101        1000500.08001Üáñ                                                   77àõÇÉÉÇ³Ù ¶»ÛÃë                  N1234567890 005 01/01/2003                             Ð        CashRequest_CashRequest_CashRequest_1 Î³ÝËÇÏ³óÙ³Ý Ñ³Ûï                                                                                      "
    
    
    Call CheckQueryRowCount("FOLDERS","fISN",CashOutIsn,2)
    Call CheckDB_FOLDERS(dbFOLDERS(7),1)
    Call CheckDB_FOLDERS(dbFOLDERS(8),1)
    
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",CashOutIsn,8)                       
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "11", "1630171","391311144.00", "001", "978277.86", "MSC", "C")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "11", "1714909","391311144.00", "001", "978277.86", "MSC", "D")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "11", "1629177","1333333.20", "000", "1333333.20", "MSC", "C")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "11", "1714909","1333333.20", "001", "0.00", "MSC", "D")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "11", "1630170","7555554.80", "000", "7555554.80", "CEX", "C")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "11", "1714909","7555554.80", "001", "22222.22", "CEX", "D")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "11", "1630170","180158.00", "000", "180158.00", "FEE", "D")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "11", "1630420","180158.00", "000", "180158.00", "FEE", "C")
    
    'SQL Ստուգում HI2 աղուսյակում 
    Call CheckQueryRowCount("HI2","fBASE",CashOutIsn,1)
    Call CheckDB_HI2(dbHI2(2),1)

    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckDB_HIREST("11", "1630171","671508758.50","001","1291284.50",1)
    Call CheckDB_HIREST("11", "1714909","400200032.00","001","1000500.08",1)
    Call CheckDB_HIREST("11", "1629177","-1742611.70","000","-1742611.70",1)    
    Call CheckDB_HIREST("11", "1630170","76089245.80","000","76089245.80",1)   
    Call CheckDB_HIREST("11", "1630170","76089245.80","000","76089245.80",1)   
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''-- "Աշխատանքային փաստաթղթեր" թղթապանակից Կանխիկ ելք փաստաթուղթը "Ուղարկել հաստատման" --'''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Աշխատանքային փաստաթղթեր թղթապանակից Կանխիկ ելք փաստաթուղթը Ուղարկել հաստատման --",,,DivideColor       
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ
    Call ChangeWorkspace(c_ChiefAcc)
    Call ToCountPayment(c_SendToVer,"020221") 
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''-- Գլխավոր հաշվապահ/Հաստատվող փաստաթղթեր(|) թղթապանակից կատարել "Վավերացնել" գործողությունը --''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կատարել Վավերացնել գործողությունը --",,,DivideColor   

    Set VerificationDoc = New_VerificationDocument()
        VerificationDoc.DocType = "KasRsOrd"
    
    Call GoToVerificationDocument("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ (I)",VerificationDoc) 
    
    If WaitForPttel("frmPttel") Then
        If SearchInPttel("frmPttel",7, "1000500.08") Then
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(c_ToConfirm)
            BuiltIn.Delay(2000)
            Call ClickCmdButton(1, "Ð³ëï³ï»É")
        Else 
            Log.Error "Տողը չի գտնվել Հաստատվող փաստաթղթեր(|) թղթապանակում" ,,,ErrorColor
        End If
        Call Close_Pttel("frmPttel")
     Else
        Log.Error "Can Not Open Հաստատվող փաստաթղթեր(|) Window",,,ErrorColor      
     End If 
     
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  ACSBRANCH:00  ACSDEPART:1  BLREP:0  OPERTYPE:MSC  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  USERID:  77  ISTLLCREATED:1  DATE:20210202  ACCDB:00057133311  CUR:001  KASSA:001  ACCCR:000001101  SUMMA:1000500.08  TOTAL:1005600.01  KASSIMV:051  BASE:Î³ÝËÇÏ³óÙ³Ý Ñ³Ûï  AIM:CashRequest_CashRequest_CashRequest_1  RECEIVER:àõÇÉÉÇ³Ù  RECEIVERLASTNAME:¶»ÛÃë  PASSNUM:N1234567890  PASBY:005  DATEPASS:20030101  DATEEXPIRE:20260101  CITIZENSHIP:1  COUNTRY:AM  ADDRESS:ºñ¨³Ý, ÐÐ, ²ëï³ýÛ³Ý 1,1  FROMPAYORD:0  ACSBRANCHINC:00  ACSDEPARTINC:1  CHRGACC:000001100  TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  CHRGCUR:000  CHRGCBCRS:1/1  PAYSCALE:01  CHRGSUM:180158  PRSNT:0.1  CHRGINC:000919400  NOCRGSUM:550105.06  FRSHNOCRG:0  CURTES:1  CURVAIR:3  VOLORT:7  NONREZ:0  JURSTAT:21  COMM:¶³ÝÓáõÙ Ï³ÝËÇÏ³óáõÙÇó  PAYSYSIN:Ð  XSUM:22222.22  XCUR:000  XACC:000001100  XDLCRS:   340.0000/    1  XDLCRSNAME:000 / 001  XCBCRS:400.0000/1  XCBCRSNAME:000 / 001  XCUPUSA:1  XCURSUM:7555554.8  XSUMMAIN:978277.86  XINC:000931900  XEXP:001434300  SYSCASE:CashReq  NOTSENDABLE:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",CashOutIsn,1)
    Call CheckDB_DOCS(CashOutIsn,"KasRsOrd","14",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",CashOutIsn,5)
    Call CheckDB_DOCLOG(CashOutIsn,"77","M","14","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Call CheckQueryRowCount("FOLDERS","fISN",CashOutIsn,0)
    
    'SQL Ստուգում PAYMENTS աղուսյակում  
    Call CheckQueryRowCount("PAYMENTS","fISN",CashOutIsn,1)
    
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",CashOutIsn,9)   
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "CE", "1578250","7555554.80", "001", "22222.22", "PUR", "D")     
                   
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "01", "1630171","391311144.00", "001", "978277.86", "MSC", "C")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "01", "1714909","391311144.00", "001", "978277.86", "MSC", "D")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "01", "1629177","1333333.20", "000", "1333333.20", "MSC", "C")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "01", "1714909","1333333.20", "001", "0.00", "MSC", "D")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "01", "1630170","7555554.80", "000", "7555554.80", "CEX", "C")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "01", "1714909","7555554.80", "001", "22222.22", "CEX", "D")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "01", "1630170","180158.00", "000", "180158.00", "FEE", "D")
    Call Check_HI_CE_accounting ("20210202",CashOutIsn, "01", "1630420","180158.00", "000", "180158.00", "FEE", "C")
    
    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckDB_HIREST("01", "1630171","4283940461.90","001","10197874.97",1)
    Call CheckDB_HIREST("11", "1630171","1062819902.50","001","2269562.36",1)
    Call CheckDB_HIREST("01", "1714909","150271340.60","001","379241.05",1)
    Call CheckDB_HIREST("01", "1629177","-22869693.80","000","-22869693.80",1)    
    Call CheckDB_HIREST("01", "1630170","1091978329.80","000","1091978329.80",1)   
    Call CheckDB_HIREST("11", "1630170","83464642.60","000","83464642.60",1)   
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''-"Հաշվառված վճարային փաստաթղթերից" հեռացնել "Կանխիք ելք" և "Կանխիկ մուտքի" գործողությունները-'''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հեռացնել Կանխիք ելք և Կանխիկ մնացորդի ճշտում գործողությունները --",,,DivideColor     
    
    wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    'Լրացնել "Ամսաթիվ" դաշտը
    Call Rekvizit_Fill("Dialog",1,"General","PERN", "030320")
    Call Rekvizit_Fill("Dialog",1,"General","PERK", "020221")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    If WaitForPttel("frmPttel") Then
        Call SearchAndDelete("frmPttel", 1, "Î³ÝËÇÏ »Éù", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        BuiltIn.Delay(2000)
        Call SearchAndDelete("frmPttel", 1, "Î³ÝËÇÏ Ùáõïù", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        BuiltIn.Delay(2000)
        Call Close_Pttel("frmPttel")
     Else
        Log.Error "Can Not Open Հաշվառված վճարային փաստաթղթեր Window",,,ErrorColor      
     End If     
     
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''-"Նոր ստեղծված կանխիկացման հայտ թղթապանակից հեռացնել "կանխիկացման հայտ" փաստաթուղթը-''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հեռացնել կանխիկացման հայտ փաստաթուղթը --",,,DivideColor     
    
    'Մուտք Հաճախորդի սպասարկում և դրամարկղ (ընդլայնված) ԱՇՏ
    Call ChangeWorkspace(c_CustomerService)
    Call GoTo_CustomerService_CashRequest(NewCashRequests)
        
    If SearchInPttel("frmPttel",8, "1005600.01") Then
        Call SearchAndDelete("frmPttel", 2, "00000075", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")   
    Else 
        Log.Error "Տողը չի գտնվել Նոր ստեղծված կանխիկացման հայտ թղթապանակում" ,,,ErrorColor
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''-- Փոխել CASHACCFLAG Պարամետրի արժեքը --''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Փոխել CASHACCFLAG Պարամետրի արժեքը --" ,,, DivideColor     
    
    'Մուտք Ադմինիստրատորի ԱՇՏ 4.0
    Call ChangeWorkspace(c_Admin40)
    
    Call GoTo_SystemParameters("CASHACCFLAG") 
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToEdit)
    wMDIClient.Refresh
    
    'Լրացնել "Արժեք" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "CheckBox", "VALUE", "0")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    BuiltIn.Delay(2000)
    Call Close_Pttel("frmPttel")    
    
    
    Call Close_AsBank()   
End Sub   

Sub SQL_Initialize_Cash_Acc_ByCashReq(fISN)

    Set dbFOLDERS(1) = New_DB_FOLDERS()
    With dbFOLDERS(1)
        .fFOLDERID = "C.1628383"
        .fNAME = "KasPrOrd"
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "5"
        .fCOM = "Î³ÝËÇÏ Ùáõïù"
        .fSPEC = "²Ùë³ÃÇí- 03/03/20 N- 000874 ¶áõÙ³ñ-           550,005.06 ²ñÅ.- 001 [Üáñ]"
        .fECOM = "Cash Deposit Advice"
    End With 

    Set dbFOLDERS(2) = New_DB_FOLDERS()
    With dbFOLDERS(2)
        .fFOLDERID = "Oper.20200303"
        .fNAME = "KasPrOrd"
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "5"
        .fCOM = "Î³ÝËÇÏ Ùáõïù"
        .fSPEC = "00087477700000001101  7770000007560101       550005.06001Üáñ                                                   77master 1_²ÝáõÝ_²ÝáõÝ_²ÝáõÝ_²ÝáõÝN1234567890 005 01/01/2003                                      Ð³Ù³Ó³ÛÝ Ã. Ñ³ßíÇ Ð³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ                                                                                                       "
        .fECOM = "Cash Deposit Advice"
        .fDCBRANCH = "00"
        .fDCDEPART = "1"
    End With  
    
    Set dbFOLDERS(3) = New_DB_FOLDERS()
    With dbFOLDERS(3)
        .fFOLDERID = "Ver.20200303001"
        .fNAME = "KasPrOrd"
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "4"
        .fCOM = "Î³ÝËÇÏ Ùáõïù"
        .fSPEC = "00087877700000001101  7770000007560101       550005.06001  77Ð³Ù³Ó³ÛÝ Ã. Ñ³ßíÇ               Ð³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ             master 1_²ÝáõÝ_²ÝáõÝ_²ÝáõÝ_²ÝáõÝ"
        .fECOM = "Cash Deposit Advice"
        .fDCBRANCH = "00"
        .fDCDEPART = "1"
    End With 
    
    Set dbFOLDERS(4) = New_DB_FOLDERS()
    With dbFOLDERS(4)
        .fFOLDERID = "C.1628383"
        .fNAME = "CBCshReq"
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "Î³ÝËÇÏ³óÙ³Ý Ñ³Ûï"
        .fSPEC = ""
        .fECOM = "Cash request"
    End With  
    Set dbFOLDERS(5) = New_DB_FOLDERS()
    With dbFOLDERS(5)
        .fFOLDERID = "CashReq.20240505"
        .fNAME = "CBCshReq"
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "Î³ÝËÇÏ³óÙ³Ý Ñ³Ûï"
        .fSPEC = "0000007500057133311      1005600.01001CashRequest_CashRequest_CashRequ0020Üáñ                   77àõÇÉÉÇ³Ù ¶»ÛÃë                  William H. Gates                àõÇÉÉÇ³Ù ¶»ÛÃë                  "
        .fECOM = "Cash request"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With  
    
    Set dbFOLDERS(6) = New_DB_FOLDERS()
    With dbFOLDERS(6)
        .fFOLDERID = "Oper.20240505"
        .fNAME = "CBCshReq"
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "Î³ÝËÇÏ³óÙ³Ý Ñ³Ûï"
        .fSPEC = "0009097770000057133311                      1005600.01001Üáñ                                                   77àõÇÉÉÇ³Ù ¶»ÛÃë                                                                         Ð        CashRequest_CashRequest_CashRequest_1                                                                                                       "
        .fECOM = "Cash request"
        .fDCBRANCH = "00"
        .fDCDEPART = "1"
    End With 
    
    Set dbFOLDERS(7) = New_DB_FOLDERS()
    With dbFOLDERS(7)
        .fFOLDERID = "C.1628383"
        .fNAME = "KasRsOrd"
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "5"
        .fCOM = "Î³ÝËÇÏ »Éù"
        .fSPEC = "²Ùë³ÃÇí- 02/02/21 N- 000890 ¶áõÙ³ñ-         1,000,500.08 ²ñÅ.- 001 [Üáñ]"
        .fECOM = "Cash Withdrawal Advice"
    End With 
    
    Set dbFOLDERS(8) = New_DB_FOLDERS()
    With dbFOLDERS(8)
        .fFOLDERID = "Oper.20210202"
        .fNAME = "KasRsOrd"
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "5"
        .fCOM = "Î³ÝËÇÏ »Éù"
        .fSPEC = "000890777000005713331177700000001101        1000500.08001Üáñ                                                   77àõÇÉÉÇ³Ù ¶»ÛÃë                  N1234567890                                            Ð        CashRequest_CashRequest_CashRequest_1 Î³ÝËÇÏ³óÙ³Ý Ñ³Ûï                                                                                      "
        .fECOM = "Cash Withdrawal Advice"
        .fDCBRANCH = "00"
        .fDCDEPART = "1"
    End With 
    
    Set dbHI2(1) = New_DB_HI2()
    With dbHI2(1)
        .fDATE = "2020-03-03"
        .fTYPE = "10"
        .fOBJECT = "1628383"
        .fGLACC = "1578250"
        .fSUM = "0.00"
        .fCUR = "001"
        .fCURSUM = "550005.06"
        .fOP = "MSC"
        .fBASE = fISN
        .fDBCR = "D"
    End With
    
    Set dbHI2(2) = New_DB_HI2()
    With dbHI2(2)
        .fDATE = "2021-02-02"
        .fTYPE = "10"
        .fOBJECT = "1628383"
        .fGLACC = "1578250"
        .fSUM = "0.00"
        .fCUR = "001"
        .fCURSUM = "550105.06"
        .fOP = "MSC"
        .fBASE = fISN
        .fDBCR = "C"
    End With
End Sub