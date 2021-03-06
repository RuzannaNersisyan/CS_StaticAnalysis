'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT Library_CheckDB
'USEUNIT Mortgage_Library
'USEUNIT Library_Contracts
'USEUNIT Card_Library
'USEUNIT Loan_Agreements_Library 
'USEUNIT Akreditiv_Library
'USEUNIT CashInput_Confirmphases_Library
'USEUNIT  Main_Accountant_Filter_Library
'USEUNIT Loan_Agreemnts_With_Schedule_Library

'USEUNIT Constants
Option Explicit

'Test Case Id - 171665

Dim dbFOLDERS(4),dbHI2(4)
    
Sub Cash_Accounting_ByCredit()
  
    Dim sDATE,eDATE
    Dim Client,Acc,Loan,FolderName,CashIn,CashInIsn
    Dim VerificationDoc,CashAccountingFilter,SummaryOfContracts
    Dim Return_Payed_Isn,GiveCreditIsn,FromProvisionIsn
    Dim fBODY
    
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    sDATE = "20030101"
    eDATE = "20260101"
    Call Initialize_AsBank("bank", sDATE, eDATE)
    Login("ARMSOFT")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-- Կարգավորումներում կատարել համապատասխան փոփոխությունները կանխիկ հաշվառման համար --''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կարգավորումներում կատարել համապատասխան փոփոխությունները կանխիկ հաշվառման համար --" ,,, DivideColor      
    
    'Մուտք գործել "Ենթահամակրգեր(ՀԾ)"
    Call ChangeWorkspace(c_Subsystems) 
    Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|²¹ÙÇÝÇëïñ³ïÇí Ù³ë|Î³ñ·³íáñáõÙÝ»ñ ¨ ¹ñáõÛÃÝ»ñ|²ÝÏ³ÝËÇÏ ·áñÍ³ñùÝ»ñÇó Ï³ÝËÇÏÇ Ñ³ßí³éÙ³Ý ¹ñáõÛÃÝ»ñ|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)")
    BuiltIn.Delay(2000)
    With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
   
      .Row = 0
      .Col = 0
      .Keys("0005")
      .Row = 0
      .Col = 4
      .Keys("Agr ")
      
      .Row = 1
      .Col = 4
      .Keys("Ret ")
      
      .Row = 2
      .Col = 4
      .Keys("RtGz")
    End With
    
    Call ClickCmdButton(1, "Î³ï³ñ»É")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''-- Հաշիվներ թղթապանակից կատարել "Ավելացնել" գործողությունը --''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հաշիվներ թղթապանակից կատարել Ավելացնել գործողությունը --" ,,, DivideColor    
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ
    Call ChangeWorkspace(c_ChiefAcc)
    Call wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³ßÇíÝ»ñ") 
    BuiltIn.Delay(1000)
    'Կանխիկ հաշվառման և Բացման ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CASHAC", 1)
    Call Rekvizit_Fill("Dialog", 1, "General", "DATOTKN", "010120")
    Call Rekvizit_Fill("Dialog", 1, "General", "DATOTKK", "010120")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)

    Set Acc = New_Account()
    With Acc         
      .BalanceAccount = "3022000"
      .AccountHolder = "00000678"
      .Name = ""
      .EnglishName = ""
      .RemainderType = ""
      .Curr = "000"
      .AccountType = "01"
      .OpenDate = "010120"
      .Account = ""
      .AccessType = "01"
      .CashAccounting = 1
    End With
    
    Call Create_Account(Acc)
    Log.Message "fISN = " & Acc.Isn ,,,SqlDivideColor
    Call SQL_Initialize_Cash_Accounting_ByCredit(Acc.Isn)
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  CLIMAINACC:0  BALACC:3022000  CLICOD:00000678  NAME:KERAMIKA Ê³ã³ïãÛ³Ý ìÉ³¹»Ý ì³ÝÇÏÇ  ENAME:Ê³ã³ïãÛ³Ý ìÉ³¹»Ý ì³ÝÇÏÇ  DK:2  CODVAL:000  ACCTYPE:01  DATOTK:20200101  CODE:00001283022  BLREP:0  ACSBRANCH:00  ACSDEPART:1  ACSTYPE:01  ULIMIT:999999999999.99  CASHAC:1  BALACC2:999999  BALACC3:999999  FROZEN:0  FNSTATE:2  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",Acc.Isn,1)
    Call CheckDB_DOCS(Acc.Isn,"Acc     ","2",fBODY,1)
    
    'SQL Ստուգում ACCOUNTS աղուսյակում
    Call CheckQueryRowCount("ACCOUNTS","fISN",Acc.Isn,1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Call CheckQueryRowCount("FOLDERS","fISN",Acc.Isn,2)
    dbFOLDERS(1).fSPEC = Acc.Account&"  ²ñÅ.- 000  îÇå- 01  Ð/Ð³ßÇí- 3022000   ²Ýí³ÝáõÙ-KERAMIKA Ê³ã³ïãÛ³Ý ìÉ³¹»Ý ì³ÝÇÏÇ"
    dbFOLDERS(2).fKEY = Acc.Account
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)
    
    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckDB_HIREST("01", Acc.Isn,"0.00","XXX","-999999999999.99",1)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''-- Հաշիվներ թղթապանակից կատարել "Կանխիկ մուտք" գործողությունը --'''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հաշիվներ թղթապանակից կատարել Կանխիկ մուտք գործողությունը --" ,,, DivideColor 
    Log.Message Acc.Account
    
    Call SearchInPttel("frmPttel",0, Acc.Account)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_InnerOpers &"|"& c_Cashin)             
      
    Set CashIn = New_CashIn()  
    With CashIn
        .Date = "010120"
        .Amount = "100000"
        .CashLabel = "022"
        .Base = "Ð³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ"
        .Aim = "Ð³Ù³Ó³ÛÝ Ã. Ñ³ßíÇ"
        .Depositor = "00000678"
        .FirstName = "master"
    End With 

    CashInIsn = Fill_CashIn(CashIn)
    
    'Եթե քաղվածքի պատուհանը հայտնվել է, ապա փակում է
    If wMDIClient.VBObject("FrmSpr").Exists Then
        wMDIClient.VBObject("FrmSpr").Close
    Else
        Log.Error "Statement window doesn't exist!",,,ErrorColor
    End If
    Call Close_Pttel("frmPttel")
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''-- "Աշխատանքային փաստաթղթեր" թղթապանակից կատարել "Ուղարկել հաստատման" գործողությունը --'''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կատարել Ուղարկել հաստատման գործողությունը --",,,DivideColor       
    
    wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
    'Լրացնել "Ամսաթիվ" դաշտերը
    Call Rekvizit_Fill("Dialog",1,"General","PERN", "010120")
    Call Rekvizit_Fill("Dialog",1,"General","PERK", "010120")
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''-- Գլխավոր հաշվապահ/Հաստատվող փաստաթղթեր(|) թղթապանակից կատարել "Վավերացնել" գործողությունը --''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կատարել Վավերացնել գործողությունը --",,,DivideColor   

    Set VerificationDoc = New_VerificationDocument()
        VerificationDoc.DocType = "KasPrOrd"
    
    Call GoToVerificationDocument("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ (I)",VerificationDoc) 
    
    If WaitForPttel("frmPttel") Then
        If SearchInPttel("frmPttel",7, "100000") Then
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
     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''-- "Աշխատանքային փաստաթղթեր" թղթապանակից կատարել "Վավերացնել" գործողությունը --'''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Աշխատանքային փաստաթղթեր թղթապանակից կատարել Վավերացնել գործողությունը --",,,DivideColor       
    
    wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
    'Լրացնել "Ամսաթիվ" դաշտերը
    Call Rekvizit_Fill("Dialog",1,"General","PERN", "010120")
    Call Rekvizit_Fill("Dialog",1,"General","PERK", "010120")
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
    
     Log.Message "fISN = " & CashInIsn ,,,SqlDivideColor
     Call SQL_Initialize_Cash_Accounting_ByCredit(CashInIsn)
       
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  ACSBRANCH:00  ACSDEPART:1  BLREP:0  OPERTYPE:MSC  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  USERID:  77  DATE:20200101  KASSA:001  ACCDB:000001100  CUR:000  ACCCR:"&Acc.Account&"  SUMMA:100000  KASSIMV:022  BASE:Ð³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ  AIM:Ð³Ù³Ó³ÛÝ Ã. Ñ³ßíÇ  CLICODE:00000678  PAYER:master  ISTLLCREATED:1  ACSBRANCHINC:00  ACSDEPARTINC:1  CHRGACC:"&Acc.Account&"  TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  CHRGCUR:000  CHRGCBCRS:1/1  VOLORT:9X  NONREZ:0  JURSTAT:11  USEOVERLIMIT:0  NOTSENDABLE:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",CashInIsn,1)
    Call CheckDB_DOCS(CashInIsn,"KasPrOrd","11",fBODY,1)
    
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",CashInIsn,2)                       
    Call Check_HI_CE_accounting ("20200101",CashInIsn, "01", "1630170","100000.00", "000", "100000.00", "MSC", "D")
    Call Check_HI_CE_accounting ("20200101",CashInIsn, "01", Acc.Isn,"100000.00", "000", "100000.00", "MSC", "C")

    'SQL Ստուգում HI2 աղուսյակում 
    Call CheckQueryRowCount("HI2","fBASE",CashInIsn,1)
    dbHI2(1).fBASE = CashInIsn
    Call CheckDB_HI2(dbHI2(1),1)

    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckQueryRowCount("HIREST","fOBJECT",Acc.Isn,10)
    Call CheckDB_HIREST("01", Acc.Isn,"-100000.00","000","-100000.00",1)
    
    'SQL Ստուգում PAYMENTS աղուսյակում  
    Call CheckQueryRowCount("PAYMENTS","fISN",CashInIsn,1)
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''-- Վարկեր (տեղաբաշխված) ԱՇՏ-ում ստեղծում ենք Գրաֆիկով վարկային պայմանագիր --''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Վարկեր (տեղաբաշխված) ԱՇՏ-ում ստեղծում ենք Գրաֆիկով վարկային պայմանագիր --" ,,, DivideColor        
    
    'Մուտք Վարկեր (տեղաբաշխված) ԱՇՏ
    Call ChangeWorkspace(c_Loans)

    FolderName = "|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|"
    Set Loan = New_LoanDocument()
    With Loan
      .CalcAcc = Acc.Account
      .Limit = "55555"
      .ArgType = "0005"
      .Date = "080821"
      .GiveDate = "080821"
      .Term = "030822"
      .FirstDate = "080821"
      .PaperCode = 123456789
      .DocType = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ"

      Call .CreatePlLoan("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
      Log.Message .DocNum

      Call Close_Pttel("frmPttel")
  
      'Պայմանագրին ուղղարկել հաստատման
      .SendToVerify(FolderName & "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
      .Verify(FolderName & "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
      
      Log.Message "fISN = " & .fBASE ,,,SqlDivideColor
      
      'SQL Ստուգում DOCS աղուսյակում
      fBODY = "  CLICOD:00000678  NAME:KERAMIKA Ê³ã³ïãÛ³Ý ìÉ³¹»Ý ì³ÝÇÏÇ  AGRTYPE:0005  CURRENCY:000  SUMMA:55555  DATE:20210808  DATEGIVE:20210808  DATEAGR:20220208  EXISTSPROLPERSCH:0  ISLINE:0  ALLOCATEWITHLIM:0  ISREGENERATIVE:0  ISCRCARD:0  AUTOCAP:0  ISLIMPERPR:0  ISPERPR:0  ACSBRANCH:00  ACSDEPART:1  ACSTYPE:C10  AUTODEBT:1  DEBTJPART:0  USECLICONNSCH:0  USECODEBTORSACCS:0  ONLYOVERDUE:0  DATESFILLTYPE:2  AGRMARBEG:20210808  AGRMARFIN:20220208  ISNOTUSETHISM:0  AGRPERIOD:1/0  PASSOVDIRECTION:2  PASSOVTYPE:0  SUMSDATESFILLTYPE:1  SUMSFILLTYPE:01  FILLROUND:2  MIXEDSUMSINSCH:0  FIXEDROWSINSCH:1  APARTPERDATES:0  KINDSCALE:1  PCAGR:12.0000/365  PCNOCHOOSE:8.0000/365  PCGRANT:0/1  CONSTPER:0  ISCONSCURPRD:0  FILLROUNDPR:2  DONOTCALCPCBASE:0  PAYPERGIVE:0  PAYPERGIVEPER:0  PCNDER:12.6822  PCNDERALL:13.1034  PCNDERAUTO:1  KINDPENCALC:1  PCPENAGR:0.1000/1  PCPENPER:0.1000/1  PCLOSS:0/1  CALCFINPER:0  CALCJOUTS:0  SECTOR:U2  USAGEFIELD:01.001  AIM:00  SCHEDULE:9  GUARANTEE:9  COUNTRY:AM  LRDISTR:001  REGION:010000008  PERRES:1  REDUCEOVRDDAYS:0  WEIGHTAMDRISK:0  PPRCODE:123456789  TIMEOP:11:11:11  CHRGFIRSTDAY:1  GIVEN:0  SUBJRISK:0  UPDINS:0  DOOVRDINWORKDAYS:0  ISNBOUT:0  PUTINLR:1  NOTCLASS:0  OTHERCOLLATERAL:0  OVRDDAYSCALCACRA:0  LASTOVRDDATEACRA:0  OVRDAGRSUMACRA:0  OVRDPERSUMACRA:0  RISKACRA:0  LASTCLASSDATEACRA:0  "
      fBODY = Replace(fBODY, "  ", "%")
      Call CheckQueryRowCount("DOCS","fISN",.fBASE,1)
      Call CheckDB_DOCS(.fBASE,"C1Univer","7",fBODY,1)
      
      'SQL Ստուգում AGRNOTES աղուսյակում
      Call CheckQueryRowCount("AGRNOTES","fAGRISN",.fBASE,1)
      'SQL Ստուգում AGRSCHEDULE աղուսյակում
      Call CheckQueryRowCount("AGRSCHEDULE","fAGRISN",.fBASE,1)
      'SQL Ստուգում AGRSCHEDULEVALUES աղուսյակում
      Call CheckQueryRowCount("AGRSCHEDULEVALUES","fAGRISN",.fBASE,12)
      'SQL Ստուգում CAGRACCS աղուսյակում
      Call CheckQueryRowCount("CAGRACCS","fAGRISN",.fBASE,1)
      'SQL Ստուգում CONTRACTS աղուսյակում
      Call CheckQueryRowCount("CONTRACTS","fDGISN",.fBASE,1)
      'SQL Ստուգում DOCP աղուսյակում
      Call CheckQueryRowCount("DOCP","fPARENTISN",.fBASE,1)
      'SQL Ստուգում DOCSG աղուսյակում
      Call CheckQueryRowCount("DOCSG","fISN",.fBASE,40)
      'SQL Ստուգում FOLDERS աղուսյակում
      Call CheckQueryRowCount("FOLDERS","fISN",.fBASE,4)
      'SQL Ստուգում HIF աղուսյակում
      Call CheckQueryRowCount("HIF","fOBJECT",.fBASE,19)
      'SQL Ստուգում RESNUMBERS աղուսյակում
      Call CheckQueryRowCount("RESNUMBERS","fISN",.fBASE,1)   
  
      If Not LetterOfCredit_Filter_Fill("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|", "", .DocNum) Then 
          Log.Error .DocNum & " Տողը չի գտնվել Վարկեր (Տեղաբաշխված) թղթապանակում",,,ErrorColor
      End If
  
      Log.Message "-- Կատարել Գանձում տրամադրումից գործողությունը --",,,DivideColor   
      Call Collect_From_Provision(.Date, "", 2, Acc.Account, FromProvisionIsn)
      
      Log.Message "fISN = " & FromProvisionIsn ,,,SqlDivideColor
      
      'SQL Ստուգում DOCS աղուսյակում
      fBODY = "  DATE:20210808  SUMMA:60  CASHORNO:2  APPLYCONNSCH:0  ACSBRANCH:00  ACSDEPART:1  ACSTYPE:C10  USERID:  77  "
      fBODY = Replace(fBODY, "  ", "%")
      Call CheckQueryRowCount("DOCS","fISN",FromProvisionIsn,1)
      Call CheckDB_DOCS(FromProvisionIsn,"C1DSPay ","5",fBODY,1)
      
      'SQL Ստուգում HI աղուսյակում
      Call CheckQueryRowCount("HI","fBASE",FromProvisionIsn,2)                       
      Call Check_HI_CE_accounting ("2021-08-08",FromProvisionIsn, "01", "230416894","60.00", "000", "60.00", "FEE", "C")
      Call Check_HI_CE_accounting ("2021-08-08",FromProvisionIsn, "01", Acc.Isn,"60.00", "000", "60.00", "FEE", "D")

      'SQL Ստուգում HIR աղուսյակում
      Call CheckQueryRowCount("HIR","fBASE",FromProvisionIsn,1)
      Call Check_HIR("2021-08-08", "R^", .fBASE, "000", "60.00", "PAY", "D")

      'SQL Ստուգում HIRREST  աղուսյակում
      Call CheckQueryRowCount("HIRREST","fOBJECT",.fBASE,1)
      Call CheckDB_HIRREST("R^",.fBASE,"60.00","2021-08-08",1)
  
      Log.Message "-- Կատարել Վարկի տրամադրում գործողությունը --",,,DivideColor    
      Call Give_Credit(.Date, "", 2, Acc.Account, GiveCreditIsn)
      
      Log.Message "fISN = " & GiveCreditIsn ,,,SqlDivideColor
      
      'SQL Ստուգում DOCS աղուսյակում
      fBODY =  "  DATE:20210808  SUMMA:55555  CASHORNO:2  COMMENT:ì³ñÏÇ ïñ³Ù³¹ñáõÙ  ACSBRANCH:00  ACSDEPART:1  ACSTYPE:C10  USERID:  77  SYSTEMTYPE:1  "
      fBODY = Replace(fBODY, "  ", "%")
      Call CheckQueryRowCount("DOCS","fISN",GiveCreditIsn,1)
      Call CheckDB_DOCS(GiveCreditIsn,"C1DSAgr ","5",fBODY,1)
      
      'SQL Ստուգում HI աղուսյակում
      Call CheckQueryRowCount("HI","fBASE",GiveCreditIsn,6)                       
      Call Check_HI_CE_accounting ("2021-08-08",GiveCreditIsn, "01",   "1629176",   "60.00", "000", "60.00", "MSC", "C")
      Call Check_HI_CE_accounting ("2021-08-08",GiveCreditIsn, "01",   "1629496",   "60.00", "000", "60.00", "MSC", "D")
      Call Check_HI_CE_accounting ("2021-08-08",GiveCreditIsn, "01", "230416894",   "60.00", "000", "60.00", "MSC", "D")
      Call Check_HI_CE_accounting ("2021-08-08",GiveCreditIsn, "01", Acc.Isn,"55555.00", "000", "55555.00", "MSC", "C")
      Call Check_HI_CE_accounting ("2021-08-08",GiveCreditIsn, "01", Acc.Isn,"55555.00", "000", "55555.00", "MSC", "C")
      
      'SQL Ստուգում HIRREST  աղուսյակում 
      Call CheckQueryRowCount("HIRREST","fOBJECT",.fBASE,3)
      Call CheckDB_HIRREST("R1",.fBASE,"55555.00","2021-08-08",1)
      Call CheckDB_HIRREST("R¾",.fBASE,"-60.00","2021-08-08",1) 
      
      'SQL Ստուգում HI2 աղուսյակում 
      Call CheckQueryRowCount("HI2","fBASE",GiveCreditIsn,1)
      dbHI2(2).fBASE = GiveCreditIsn
      Call CheckDB_HI2(dbHI2(2),1)
      
      'SQL Ստուգում HIF  աղուսյակում 
      Call CheckQueryRowCount("HIF","fOBJECT",.fBASE,20)
      Call Check_HIF("2021-08-08", "N0", .fBASE, "0.00", "0.00", "PNE", Null)
      
      'SQL Ստուգում HIR աղուսյակում 
      Call CheckQueryRowCount("HIR","fOBJECT",.fBASE,3)
      Call Check_HIR("2021-08-08", "R1", .fBASE, "000", "55555.00", "AGR", "D")
      Call Check_HIR("2021-08-08", "R¾", .fBASE, "000", "60.00", "PAY", "C")
      
      'SQL Ստուգում HIRREST  աղուսյակում 
      Call CheckQueryRowCount("HIRREST","fOBJECT",.fBASE,3)
      Call CheckDB_HIRREST("R1",.fBASE,"55555.00","2021-08-08",1)
      Call CheckDB_HIRREST("R^",.fBASE,"60.00","2021-08-08",1)
      Call CheckDB_HIRREST("R¾",.fBASE,"-60.00","2021-08-08",1)
      
      Call CompareFieldValue("frmPttel", "FKEY", .DocNum)
      Call CompareFieldValue("frmPttel", "FAGRREM", "55,555.00")
      
      'SQL Ստուգում HIREST  աղուսյակում  
      Call CheckDB_HIREST("01", Acc.Isn,"-155495.00","000","-155495.00",1)

      Call Close_Pttel("frmPttel") 
    End With
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''-- Կանխիկ միջոցների հաշվառում թղթապանակում ստուգել Մնացորդը --''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կանխիկ միջոցների հաշվառում թղթապանակում ստուգել Մնացորդը --" ,,, DivideColor 
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ
    Call ChangeWorkspace(c_ChiefAcc)
    
    Set CashAccountingFilter = New_CashAccounting()
    With CashAccountingFilter
      .ClientCode = "00000678"
      .Curr = "000"
    End With
    Call GoTo_CashAccounting(CashAccountingFilter)     
    
    Call CompareFieldValue("frmPttel", "FKEY", "00000678")
    Call CompareFieldValue("frmPttel", "FCUR", "000")   
    Call CompareFieldValue("frmPttel", "SUM", "155,555.00")   
    Call CompareFieldValue("frmPttel", "FCOM", "KERAMIKA Ê³ã³ïãÛ³Ý ìÉ³¹»Ý ì³ÝÇÏÇ")
    Call Close_Pttel("frmPttel")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''-- Հեռացնել Գանձում տրամադրումից և Վարկի տրամադրում գործողությունները--''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հեռացնել Գանձում տրամադրումից և Վարկի տրամադրում գործողությունները--" ,,, DivideColor     
    
    'Մուտք Վարկեր (տեղաբաշխված) ԱՇՏ
    Call ChangeWorkspace(c_Loans)    
    
    If LetterOfCredit_Filter_Fill("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|", "", Loan.DocNum) Then 
        Call Delete_Actions_ByCount("010120","010122",True,Null,c_OpersView,0)
    Else    
        Log.Error .DocNum & " Տողը չի գտնվել Վարկեր (Տեղաբաշխված) թղթապանակում",,,ErrorColor
    End If
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''-- Հեռացնել ստեղծված վարկային պայմանագիրը --'''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հեռացնել ստեղծված վարկային պայմանագիրը --" ,,, DivideColor     
   
    Call SearchAndDelete("frmPttel", 0, Loan.DocNum, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
    BuiltIn.Delay(2000)
    Call Close_Pttel("frmPttel")    
    
    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckDB_HIREST("01", Acc.Isn,"-100000.00","000","-100000.00",1)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-- Մուտք գործել Պայմանագրերի ամփոփում թղթապանակ --''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Մուտք գործել Պայմանագրերի ամփոփում թղթապանակ --" ,,, DivideColor 
    
    Set SummaryOfContracts = New_Deposites_SumOfContracts()
		With SummaryOfContracts
        .general.agreeN = "ST-003"
				.general.closeAgreeExists = true
				.general.repayDateExists = true
				.general.date = "01/01/21"
				.general.agreeLevel = "1"
				.general.showClosed = 1
				.general.showNotAllClosed = 1
				.general.amountsWithoutOverPart = 1
        .show.clientMainData = 1
				.show.agreeMainData = 1
				.show.mainAmounts = 1
				.show.mainDate = 1
				.show.accMainData = 1
				.show.overlimitAmounts = 1
				.show.notes = 1
				.show.riskyInformation = 1
				.show.clientOtherData = 1
				.show.agreeOtherData = 1
				.show.otherAmounts = 1 
				.show.otherDates = 1
				.show.addData = 1
				.show.writhdrawnAmounts = 1
				.show.depositeData = 1
				.show.addAmounts = 1 
				.show.addDates = 1
		End with

    Call GoTo_DepositesSumOfCont(folderName, "ä³ÛÙ³Ý³·ñ»ñÇ ³Ù÷á÷áõÙ", SummaryOfContracts)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-- Կատարել "Կանխավ վճարված տոկոսների վերադարձ" գործողությունը --''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կատարել Կանխավ վճարված տոկոսների վերադարձ գործողությունը --" ,,, DivideColor 
    
    Call Return_Payed_Percents("030321", "37", "2", Acc.Account, Return_Payed_Isn, false)
    
    Log.Message "fISN = " & Return_Payed_Isn ,,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY =  "  DATE:20210303  SUMMA:37  SUMPAY:37  CASHORNO:2  ISPUSA:0  COMMENT:Î³ÝË³í í×³ñí³Í ïáÏáëÝ»ñÇ í»ñ³¹³ñÓ  ACSBRANCH:00  ACSDEPART:1  ACSTYPE:C10  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",Return_Payed_Isn,1)
    Call CheckDB_DOCS(Return_Payed_Isn,"C1DSRet  ","5",fBODY,1)
      
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",Return_Payed_Isn,2)                       
    Call Check_HI_CE_accounting ("2021-03-03",Return_Payed_Isn, "01",   Acc.Isn,   "37.00", "000", "37.00", "MSC", "C")
    Call Check_HI_CE_accounting ("2021-03-03",Return_Payed_Isn, "01",   "510296648",   "37.00", "000", "37.00", "MSC", "D")   
      
    'SQL Ստուգում HI2 աղուսյակում 
    Call CheckQueryRowCount("HI2","fBASE",Return_Payed_Isn,1)
    dbHI2(3).fBASE = Return_Payed_Isn
    Call CheckDB_HI2(dbHI2(3),1)
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fBASE",Return_Payed_Isn,1)
    Call Check_HIR("2021-03-03", "R2", "131492060", "000", "37.00", "RET", "D")
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT","131492060",4)
    Call CheckDB_HIRREST("R2","131492060","-0.70","2021-03-03",1)
    
    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckDB_HIREST("01", Acc.Isn,"-100037.00","000","-100037.00",1)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-- Կատարել "Կանխավ վճարված Վարձավճարի վերադարձ" գործողությունը --''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կատարել Կանխավ վճարված Վարձավճարի վերադարձ գործողությունը --" ,,, DivideColor 
    
    Call Return_Payed_Rent("030321", "1500","2",Acc.Account,Return_Payed_Isn)
    Call Close_Pttel("frmPttel")
    
    Log.Message "fISN = " & Return_Payed_Isn ,,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY =  "  CODE:ST-003  DATE:20210303  SUMMA:1500  CASHORNO:2  ISPUSA:0  ACCCORR:"&Acc.Account&"  COMMENT:Î³ÝË³í í×³ñí³Í í³ñÓ³í×³ñÇ í»ñ³¹³ñÓ  ACSBRANCH:00  ACSDEPART:1  ACSTYPE:C10  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",Return_Payed_Isn,1)
    Call CheckDB_DOCS(Return_Payed_Isn,"C1DSRtGz","5",fBODY,1)
    
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",Return_Payed_Isn,2)                       
    Call Check_HI_CE_accounting ("2021-03-03",Return_Payed_Isn, "01",   Acc.Isn,   "1500.00", "000", "1500.00", "MSC", "C")
    Call Check_HI_CE_accounting ("2021-03-03",Return_Payed_Isn, "01",   "230416894",   "1500.00", "000", "1500.00", "MSC", "D")   
    
    'SQL Ստուգում HI2 աղուսյակում 
    Call CheckQueryRowCount("HI2","fBASE",Return_Payed_Isn,1)
    dbHI2(4).fBASE = Return_Payed_Isn
    Call CheckDB_HI2(dbHI2(4),1)
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fBASE",Return_Payed_Isn,1)
    Call Check_HIR("2021-03-03", "R@", "131492060", "000", "1500.00", "RET", "D")
    
    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckDB_HIREST("01", Acc.Isn,"-101537.00","000","-101537.00",1)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''-- Կանխիկ միջոցների հաշվառում թղթապանակում ստուգել Մնացորդը --''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կանխիկ միջոցների հաշվառում թղթապանակում ստուգել Մնացորդը --" ,,, DivideColor 
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ
    Call ChangeWorkspace(c_ChiefAcc)
    
    Set CashAccountingFilter = New_CashAccounting()
    With CashAccountingFilter
      .ClientCode = "00000678"
      .Curr = "000"
    End With
    Call GoTo_CashAccounting(CashAccountingFilter)     
    
    Call CompareFieldValue("frmPttel", "FKEY", "00000678")
    Call CompareFieldValue("frmPttel", "FCUR", "000")   
    Call CompareFieldValue("frmPttel", "SUM", "101,537.00")   
    Call CompareFieldValue("frmPttel", "FCOM", "KERAMIKA Ê³ã³ïãÛ³Ý ìÉ³¹»Ý ì³ÝÇÏÇ")
    Call Close_Pttel("frmPttel") 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-- Մուտք գործել Պայմանագրերի ամփոփում թղթապանակ --''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Մուտք գործել Պայմանագրերի ամփոփում թղթապանակ --" ,,, DivideColor 
    
    'Մուտք Վարկեր (տեղաբաշխված) ԱՇՏ
    Call ChangeWorkspace(c_Loans)   
    
    Set SummaryOfContracts = New_Deposites_SumOfContracts()
		With SummaryOfContracts
        .general.agreeN = "ST-003"
				.general.date = "01/01/21"
		End with
    Call GoTo_DepositesSumOfCont(folderName, "ä³ÛÙ³Ý³·ñ»ñÇ ³Ù÷á÷áõÙ", SummaryOfContracts)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''--Հեռացնել "Կանխավ վճարված տոկոսների վերադարձ" և "Կանխավ վճարված Վարձավճարի վերադարձ" գործողությունները--'''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հեռացնել Կանխավ վճարված տոկոսների վերադարձ և Կանխավ վճարված Վարձավճարի վերադարձ գործողությունները--" ,,, DivideColor     
     
    Call Delete_Actions_ByCount("010120","010122",True,Null,c_OpersView,0)    
    Call Close_Pttel("frmPttel")
    
    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckDB_HIREST("01", Acc.Isn,"-100000.00","000","-100000.00",1)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-- Մուտք գործել Կանխիկ միջոցների հաշվառում թղթապանակ --''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Մուտք գործել Կանխիկ միջոցների հաշվառում թղթապանակ --" ,,, DivideColor 
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ
    Call ChangeWorkspace(c_ChiefAcc)
    
    Set CashAccountingFilter = New_CashAccounting()
    With CashAccountingFilter
      .ClientCode = "00000678"
      .Curr = "000"
    End With
    Call GoTo_CashAccounting(CashAccountingFilter)     
    
    'Ստուգում "Մնացորդ", "Ժամկետանց գումարի տույժ", "Հաշվի Մնացորդ" սյուների արժեքները
    Call CompareFieldValue("frmPttel", "FKEY", "00000678")
    Call CompareFieldValue("frmPttel", "FCUR", "000")   
    Call CompareFieldValue("frmPttel", "SUM", "100,000.00")   
    Call CompareFieldValue("frmPttel", "FCOM", "KERAMIKA Ê³ã³ïãÛ³Ý ìÉ³¹»Ý ì³ÝÇÏÇ")
    Call Close_Pttel("frmPttel") 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''-"Հաշվառված վճարային փաստաթղթերից" հեռացնել "Կանխիք մուտք" գործողությունները-'''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--Հեռացնել Կանխիք մուտք գործողություննը --",,,DivideColor     
    
    wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    'Լրացնել "Ամսաթիվ" դաշտը
    Call Rekvizit_Fill("Dialog",1,"General","PERN", "010120")
    Call Rekvizit_Fill("Dialog",1,"General","PERK", "020220")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    If WaitForPttel("frmPttel") Then
        Call SearchAndDelete("frmPttel", 1, "Î³ÝËÇÏ Ùáõïù", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        BuiltIn.Delay(2000)
        Call Close_Pttel("frmPttel")
     Else
        Log.Error "Can Not Open Հաշվառված վճարային փաստաթղթեր Window",,,ErrorColor      
     End If     
     
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''-- "Հաշիվներ" թղթապանակից հեռացնել ստաղծված հաշիվը --''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հաշիվներ թղթապանակից հեռացնել ստաղծված հաշիվը --",,,DivideColor     
    
    Call wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³ßÇíÝ»ñ") 
    BuiltIn.Delay(1000)
    'Կանխիկ հաշվառման և Բացման ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CASHAC", 1)
    Call Rekvizit_Fill("Dialog", 1, "General", "DATOTKN", "010120" &"[Tab]"& "010120")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
    
    If WaitForPttel("frmPttel") Then
        Call SearchAndDelete("frmPttel", 1, Acc.Account, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        BuiltIn.Delay(2000)
        Call Close_Pttel("frmPttel")
     Else
        Log.Error "Can Not Open Հաշիվներ Window",,,ErrorColor      
     End If   
     
    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckQueryRowCount("HIREST","fOBJECT",Acc.Isn,0)
      
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-- Կարգավորումներից հանել փոփոխությունները --''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կարգավորումներից հանել փոփոխությունները --" ,,, DivideColor         
    
    'Մուտք գործել "Ենթահամակրգեր(ՀԾ)"
    Call ChangeWorkspace(c_Subsystems) 
    Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|²¹ÙÇÝÇëïñ³ïÇí Ù³ë|Î³ñ·³íáñáõÙÝ»ñ ¨ ¹ñáõÛÃÝ»ñ|²ÝÏ³ÝËÇÏ ·áñÍ³ñùÝ»ñÇó Ï³ÝËÇÏÇ Ñ³ßí³éÙ³Ý ¹ñáõÛÃÝ»ñ|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)")
    BuiltIn.Delay(2000)
    With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
   
      .Row = 2
      .Col = 4
      .Keys("^A[Del]")
      .Row = 1
      .Col = 4
      .Keys("^A[Del]")
      
      BuiltIn.Delay(2000)
      .Row = 0
      .Col = 4
      .Keys("^A[Del]")
      .Row = 0
      .Col = 0
      .Keys("^A[Del]")
    End With
    
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    Call Close_AsBank()  
End Sub 

Sub SQL_Initialize_Cash_Accounting_ByCredit(fISN)

    Set dbFOLDERS(1) = New_DB_FOLDERS()
    With dbFOLDERS(1)
        .fFOLDERID = "C.737994605"
        .fNAME = "Acc     "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "  Ð³ßÇí"
        .fECOM = "  Account"
    End With 

    Set dbFOLDERS(2) = New_DB_FOLDERS()
    With dbFOLDERS(2)
        .fFOLDERID = "FNAccStatus"
        .fNAME = "Acc     "
        .fKEY = "00001513022"
        .fISN = fISN
        .fSTATUS = "0"
        .fCOM = "Ð³ßÇí"
        .fSPEC = "0000067820200101000000002 00000000 00000000                                                                      00 1  01  01         3\30\302\3022\30220\3022000"
        .fECOM = "Account"
    End With  
    
    Set dbHI2(1) = New_DB_HI2()
    With dbHI2(1)
        .fDATE = "2020-01-01"
        .fTYPE = "10"
        .fOBJECT = "737994605"
        .fGLACC = "1559631"
        .fSUM = "0.00"
        .fCUR = "000"
        .fCURSUM = "100000.00"
        .fOP = "MSC"
        .fBASE = fISN
        .fDBCR = "D"
    End With
    
    Set dbHI2(2) = New_DB_HI2()
    With dbHI2(2)
        .fDATE = "2021-08-08"
        .fTYPE = "10"
        .fOBJECT = "737994605"
        .fGLACC = "1559631"
        .fSUM = "0.00"
        .fCUR = "000"
        .fCURSUM = "55555.00"
        .fOP = "MSC"
        .fBASE = fISN
        .fDBCR = "D"
    End With
    Set dbHI2(3) = New_DB_HI2()
    With dbHI2(3)
        .fDATE = "2021-03-03"
        .fTYPE = "10"
        .fOBJECT = "737994605"
        .fGLACC = "1559631"
        .fSUM = "0.00"
        .fCUR = "000"
        .fCURSUM = "37.00"
        .fOP = "MSC"
        .fBASE = fISN
        .fDBCR = "D"
    End With  
    Set dbHI2(4) = New_DB_HI2()
    With dbHI2(4)
        .fDATE = "2021-03-03"
        .fTYPE = "10"
        .fOBJECT = "737994605"
        .fGLACC = "1559631"
        .fSUM = "0.00"
        .fCUR = "000"
        .fCURSUM = "1500.00"
        .fOP = "MSC"
        .fBASE = fISN
        .fDBCR = "D"
    End With   
End Sub