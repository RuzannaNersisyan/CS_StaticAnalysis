'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Library_Common
'USEUNIT Constants

'------------------------------------------------------------------------------
'Ակրեդիտիվ տեսակի պայմանագրի լրացում : Ֆունկցիան վերադարձնում է պայմանագրի
'համարը և ISN-ը:
'-------------------------------------------------------------------------------

'1 fBASE  - Պայմանագրի ISN
'2 docNumber - "Պայմանագրի N"  դաշտի արժեք
'3 lientCode - "Հաճախորդ" դաշտի արժեք
'4 curr - "Արժույթ" դաշտի արժեք
'5 accacc - "Հաշվարկային հաշիվ" դաշտի արժեք
'6 summ - "Գումար" դաշտի արժեք
'7 restore - "Վերականգնվող" դաշտի արժեք
'8 dategive - "Կնքման ամսաթիվ" դաշտի արժեք
'9 date_arg - "Մարման ժամկետ" դաշտի արժեք
'10 agrIntRate - "Ակրեդիտիվի տոկոսադրույք" դաշտի արժեք
'11 agrIntRatePart - "Բած.`"  դաշըի արժեք
'12 sector - "Ճյուղայնություն" դաշտի արժեք
'13 aim - "Նպատակ" դաշտի արժեք
'13 schedule - "Ծրագիր" դաշտի արժեք
'14 guarante - "Երաշխավորություն" դաշտի արժեք
'15 district - "Մարզ"  դածտի արժեք
'16 paperCode - "Պայմ. թղթային N" դաշտի արժեք

Sub Letter_Of_Credit_Doc_Fill(fBASE, docNumber, clientCode, curr, accacc, summ, _
                              restore, dategive, date_arg , agrIntRate, agrIntRatePart, _
                              sector,aim, schedule,country, guarante, district,region, paperCode)
    
    wMDIClient.refresh
    fBASE = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Պայմանագրի համարի վերագրում փոփոխականին
    docNumber = Get_Rekvizit_Value("Document",1,"General","CODE")
    'Հաճախորդի կոդ դաշտի լրացում
    Call Rekvizit_Fill("Document",1,"General","CLICOD",clientCode)
    'Արժույթ դաշտի լրացում
    Call Rekvizit_Fill("Document",1,"General","CURRENCY",curr)
    'Հաշվարկային հաշիվ դաշտի լրացում
    Call Rekvizit_Fill("Document",1,"General","ACCACC",accacc)
    'Գումար դաշտի լրացում
    Call Rekvizit_Fill("Document",1,"General","SUMMA",summ)
    'Վերականգնվող դաշտի լրացում
    If restore Then
        Call Rekvizit_Fill("Document",1,"CheckBox","ISREGENERATIVE",1)
    End If
    'Կնքման ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document",1,"General","DATE","^A[Del]" & dategive)
    'Մարման ժամկետ դաշտի լրացում
    Call Rekvizit_Fill("Document",1,"General","DATEAGR","^A[Del]" & date_arg)
    'Ակրդիտիվի տոկոսադրույք դաշտի լրացում
    Call Rekvizit_Fill("Document",3,"General","PCAGR",agrIntRate)
    'Բաժ. դաշտի լրացում
    Call Rekvizit_Fill("Document",3,"General","PCAGR",agrIntRatePart)
    'Ճուղայնություն  դաշտի լրացում 
    Call Rekvizit_Fill("Document",6,"General","SECTOR",sector) 
    'Նպատակ դաշտի լրացում
    Call Rekvizit_Fill("Document",6,"General","AIM",aim)
    'Շրագիր դաշտի լրացում
    Call Rekvizit_Fill("Document",6,"General","SCHEDULE",schedule)
    'Երկիր դաշտի լրացում
    Call Rekvizit_Fill("Document",6,"General","COUNTRY",country)   
    'Երաշխավորություն դաշտի լրացում
    Call Rekvizit_Fill("Document",6,"General","GUARANTEE",guarante)
    'Մարզ դաշտի լրացում
    Call Rekvizit_Fill("Document",6,"General","LRDISTR",district)
    'Մարզ(ՎՌ) դաշտի լրացում
    Call Rekvizit_Fill("Document",6,"General","REGION",region)
    'Պայմ. թղթային համար դաշտի լրացում
    Call Rekvizit_Fill("Document",6,"General","PPRCODE",paperCode)
    'Կատարել կոճակի սեղմում
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    
End Sub

'--------------------------------------------------------------------------------------
'Ակրեդիտիվ տեսակի պայմանագրի առկայության ստուգում հաստատվաղ փաստաթղթեր 1 թղթապանակում :
'Ֆունկցիան վերդարձնում է true, եթե պայմանագիրը առկա է , false` եթե այն բացակայում է :
'--------------------------------------------------------------------------------------
'docNum - Փաստաթղթի համար

Function Verify_Letter_OfCredit (docNum)
    Dim my_vbObj , is_exists
    is_exists = False
    Call wTreeView.DblClickItem("|²Ïñ»¹ÇïÇí|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    BuiltIn.Delay(delay_middle)
    Set my_vbObj = wMDIClient.WaitVBObject("frmPttel", delay_middle)
    If my_vbObj.Exists Then
        Do Until Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmPttel").vbObject("tdbgView").EOF
            If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(2).Text) = Trim(docNum) Then
                is_exists = True
                Exit Do
            Else
                Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
            End If
        Loop
    Else
        Log.Message("The double input frmPttel does't exist")
    End If
    Verify_Letter_OfCredit = is_exists
    
End Function

'--------------------------------------------------------------------------------------
'"Ակցեպտավում" փաստաթղթի լրացում:Ֆունկցիան վերադարձնում է "Ակցեպտավորում"
'պայմանագրի համարը և ISN-ը :
'--------------------------------------------------------------------------------------
'accDocNum - Պայամանագռրի N դաշտի արժեք
'accDocBase - Ակցեպտավորման փաստաթղթի ISN
'accDocDate - "Ամսաթիվ" դաշտի արժեք
'accDocSumma - "Գումար" դաշտի արժեք
'accReqDate _ "Պահանջի ժամկետ " դաշտի արժեք
'respEndDate - "Պարտավորության ժամկետ" դաշտի արժեք
'respPercent - "Պարտավորության տոկոսադրույք " դաշտի արժեք
'respPercentPart - "Պարտավորության տոկոսադրույքի բաժ.` " դաշտի արժեք
'respBank - "Ծանուցող բանկ" դաշտի արժեք

Sub Acceptance_Doc_Fill(accDocNum, accDocBase, accDocDate, accDocSumma, accReqDate, _
                        respEndDate, respPercent , respPercentPart, respBank )
    Dim Str, wTabFrame_3 , wTabFrame_5, wTabStrip
    
    Call wMainForm.MainMenu.Click(c_AllActions) 
    Call wMainForm.PopupMenu.Click(c_Acceptance)
    
    fBASE = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Պայմանագրի համարի վերագրում փոփոխականին
    Str = GetVBObject ("CODE", wMDIClient.vbObject("frmASDocForm"))
    'docNumber = wMDIClient.vbObject("frmASDocForm").vbOvbject("TabFrame").vbObject(Str).VBObject("TDBMask").Text 
    ' Ամսաթիվ դաշտի լրացում
    Str = GetVBObject ("DATE", wMDIClient.vbObject("frmASDocForm"))
    wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(Str).Keys(accDocDate & "[Tab]")
    'Գումար դաշտի լրացում
    Str = GetVBObject ("SUMMA", wMDIClient.vbObject("frmASDocForm"))
    wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(Str).Keys(accDocSumma & "[Tab]")
    'Պահանջի ժամկետ դաշտի լրացում
    Str = GetVBObject ("CDATEAGR", wMDIClient.vbObject("frmASDocForm"))
    wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(Str).Keys(accReqDate & "[Tab]")
    'Պարտավորության ժամկետ դաշտի լրացում
    Str = GetVBObject ("DDATEAGR", wMDIClient.vbObject("frmASDocForm"))
    wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(Str).Keys(respEndDate & "[Tab]")
    'Պարտավորության տոկոսադրույք դաշտի լրացում
    Str = GetVBObject ("DPCAGR", wMDIClient.vbObject("frmASDocForm"))
    wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(Str).Keys(respPercent & "[Tab]")
    'Պարտավորության տոկոսադրույքի բաժ` դաշտի լրացում
    Sys.Process("AsBank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmASDocForm").vbObject("TabFrame").vbObject("AsCourse_2").vbObject("TDBNumber2").Keys(respPercentPart & "[Tab]")
    'Պարտավորության տոկոսադրույք դաշտի լրացում
    Str = GetVBObject ("CLIBANK", wMDIClient.vbObject("frmASDocForm"))
    wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(Str).Keys(respBank & "[Tab]")
    'Կատարել կոճակի սեղմում
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    
End Sub
'--------------------------------------------------------------------------------------
'Պայմանագրեր թղթապանակում փաստատթղթի առկայության ստուգում : Ֆունկցիան վերադարձնում է
'True , եթե պայմանագիրը առկա է և false, եթե այն բացակայում է :
'--------------------------------------------------------------------------------------
'docType - Պայմանագրի մակարդակ դաշտի արժեք
'docNum - Պայմանագրի N դաշտի արժեք
'FolderName - ԱՇՏ-ի անուն

Function LetterOfCredit_Filter_Fill(FolderName, docType, docNum)
    
    Dim isExists
    isExists = True
    Call wTreeView.DblClickItem(FolderName & "ä³ÛÙ³Ý³·ñ»ñ")
    BuiltIn.Delay(2000) 

    'Լրացնում է պայմանագրի մակարդակ դաշտը
    Call Rekvizit_Fill("Dialog",1,"General","LEVEL",docType)
    'Լրացնում է պայմանագրի համար դաշտը
    Call Rekvizit_Fill("Dialog",1,"General","NUM", "^A[Del]" & docNum)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(2000) 
  
    wMDIClient.Refresh
    If wMDIClient.vbObject("frmPttel").vbObject("tdbgView").VisibleRows <> 1 Then
        Log.Message "There are no document with specified ID" 
        isExists = False
    End If
    LetterOfCredit_Filter_Fill = isExists
    
End Function

'______________________________________________________________________________________________________________________________________________
' Պայմանագրերի ցուցակում փնտրում է տրված անվամբ գլխավոր պայմանագիրը
'______________________________________________________________________________________________________________________________________________
'Name - Պայմանագրի անունը

Function Find_Main_Contract(Name)
   Dim isExists,my_vbObj
    
   isExists = False
   wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Click()
 
    Set wMainForm = Sys.Process("Asbank").vbObject("MainForm")
    Set wMDIClient = wMainForm.Window("MDIClient")
    BuiltIn.Delay(delay_middle)
    
    Set my_vbObj = wMDIClient.WaitVBObject("frmPttel", delay_middle)
    If my_vbObj.Exists Then
    'Ցուցակի մեջ պնտրում է գլխավոր պայմանագիրը ,չգտնելու դեպքում է դուրս է բերում սխալ
    wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveFirst
        Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
            If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(0).Text) = Trim(Name) Then
                is_exists = True
                Exit Do
            Else
                Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
            End If
        Loop
    Else
        Log.Message("The double input frmPttel does't exist")
    End If
    Find_Main_Contract = is_exists

End Function 

'____________________________________________________________________________________________________
'Ենթահամակարգերի «Սև ցուցակ» հաստատվող փաստաթղթեր թղթապանաակ փնտրում է պայմանագիրը
'_____________________________________________________________________________________________________
'DocNum - Պայմանաագրի համարը

Function Find_Doc(DocNum)
    Dim isexists
    isexists = false
    
    'Լրացնում է պայմանագրի համարը դաշտը
    Call Rekvizit_Fill("Dialog",1,"General","NUM",DocNum)
    'Սեղմել Կատարել կոճակը
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    Set wMainForm = Sys.Process("Asbank").vbObject("MainForm")
    'Փնտրում է պայմանագիրը ցուցակում, չգտնելու դեպքում դուրս է բերում սխալ
      If Trim(wMainForm.Window("MDIClient", "", 1).vbObject("frmPttel").vbObject("tdbgView").Columns.Item(2).Text) = Trim(DocNum) Then
        isexists = True     
    Else
        Log.Error("The  document  does't exist")
    End If          
    Find_Doc = isexists
  
End Function

'______________________________________________________________________________________________________________________________
'Վավերացնում է գլխավոր պայմանագիրը 
'______________________________________________________________________________________________________________________________

Sub Validate_Doc()
   BuiltIn. Delay(3000)
   Call wMainForm.MainMenu.Click(c_AllActions)
   Call wMainForm.PopupMenu.Click(c_ToConfirm)
   Call ClickCmdButton(1, "Ð³ëï³ï»É")
   BuiltIn.Delay(1000)
End Sub 

'________________________________________________________________________________________________________________________________________________
'Ստուգում է պայմանագրի առկայությունը Ակրեդիտիվ|Պայմանագրեր թղթապանակում
'________________________________________________________________________________________________________________________________________________
'Doc_Level - Պայմանագրի մակարդակ
'DocNum - Պայմանագրի համար

Function Goto_Doc_Check(Doc_Level,DocNum)

  Dim isExists
  isExists = false
 
  'Լրացնում է պայմանագրի մակարդակ դաշտը
  Call Rekvizit_Fill("Dialog",1,"General","LEVEL",Doc_Level)
  'Լրացնում է պայմանագրի համար դաշտը
  Call Rekvizit_Fill("Dialog",1,"General","NUM",DocNum)
  Call ClickCmdButton(2, "Î³ï³ñ»É")
      'Կատարում է ստուգում , եթե գտնում է պայմանագիրը վերադարձնում է TRUE ,հակառակ դեպքում դուրս է բերում սխալ
  If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(0).Text) = Trim(DocNum) Then
      isExists = True
  Else
      Log.Error("The  document  does't exist")
  End If
  Goto_Doc_Check = isExists
  
End Function 


'________________________________________________________________________________________________________________________________________
'Ստուգում է , որ Սահմանաչափեր թղթապանակում  սահմանաչապը հավասար լինի մայր գումարին
'________________________________________________________________________________________________________________________________________
'StartDate - Սկզբնաժամկետ
'EndDate - Վերջնաժամկետ
'TF - Միայն փոփոխություններ դաշտ
'DLimit - Սահմանաչափ

Function Check_Limit(StartDate,EndDate, TF, DLimit)
  
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_ViewEdit & "|" & c_Other & "|" & c_Limits)
  'Լրացնում է սկզբնաժամկետ դաշտը
  Call Rekvizit_Fill("Dialog",1,"General","START","![End]" & "[Del]" & StartDate)
  'Լրացնում է վերջնաժամկետ դաշտը
  Call Rekvizit_Fill("Dialog",1,"General","END","![End]" & "[Del]" & EndDate)
  'Լրացնում է Միայն փոփոխություններ դաշտը
  Call Rekvizit_Fill("Dialog",1,"CheckBox","ONLYCH",TF)
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  'Ստուգում է, որ սահմանաչափը հավասար լինի մայր գումարին , հակառակ դեպքւմ դուրս կբերի սխալ
  isExists = False
  wMDIClient.Refresh
  If Trim(wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").Columns.Item(3).Text) = Trim(DLimit) Then
      isExists = True
  End If
  Check_Limit = isExists
  wMDIClient.VBObject("frmPttel_2").Close

End Function 
 


'________________________________________________________________________________________________________________________________________________
'Ստուգում է ,որ Ակցեպտավորում թղթապանակում պայմանագիրը լինի միակը
'________________________________________________________________________________________________________________________________________________

Function Chek_One_Doc()
  Dim count,isOne
  isOne = false
   
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_Acceptances)

  'Կատարում է ստուգում ,եթե միանկն է վերադարձնում է TRUE ,հակռակ դեպքում դուրս է բերում սխալ
  If Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").ApproxCount= 1  then 
      isOne = True
      Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).Refresh
      Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_2").Close()
  Else
      Log.Error("There is more than one document")
  End If
  Chek_One_Doc = isOne
  
End Function 
 

'______________________________________________________________________________________________________________________________________________________________________________
'Տոկոսների հաշվարկման փաստաթղթի ստեղծում
'______________________________________________________________________________________________________________________________________________________________________________
'Calculate_Date  - հաշվարկման ամսաթիվ
'Action_Date - գործողության ամսաթիվ
'fBase - Տոկոսների հաշվարկման փաստաթղթի ISN

Sub Calculate_Percent(fBase, Calculate_Date , Action_Date)

  BuiltIn.Delay(2500)
  wMDIClient.Refresh
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_Interests & "|" & c_PrcAccruing)
  fBase = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
  
  'Լրացնում է հաշվարկման ամսաթիվ դաշտը
  Call Rekvizit_Fill("Document",1,"General","DATECHARGE","![End]" & "[Del]" & Calculate_Date )
  'Լրացնում է գործողության ամսաթիվ դաշտը
  Call Rekvizit_Fill("Document",1,"General","DATE","![End]" & "[Del]" & Action_Date)
  Call ClickCmdButton(1, "Î³ï³ñ»É")
   
End Sub 



'________________________________________________________________________________________________________________________
'Ընդհանուր դիտում
'________________________________________________________________________________________________________________________
'CDate - Ամսաթիվ

Sub Watch_Contract(CDate)
   
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_References & "|" & c_CommView)
  'Լրացնել Ամսաթիվ դաշտը
  Call Rekvizit_Fill("Dialog",1,"General","LASTDATE",CDate )
'  Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate").Keys(CDate  & "[Tab]")
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  wMDIClient.Refresh
  'Կատարում է ստուգում, եթե պատուհանը հայտնվել է ,ապա փակում է ,հակառակ դեպքում դուրս է բերում սխալ
  If wMDIClient.VBObject("FrmSpr").Exists Then
      wMDIClient.VBObject("FrmSpr").Close
  Else
      Log.error ("The window doesn't exist")
  End If
  
End Sub 
 
'_________________________________________________________________________________________________________________________________________________________________________
'Ստուգում է ,որ սահմանաչափը նվազած լինի մարված գումարի չափով
'_________________________________________________________________________________________________________________________________________________________________________
'StartDate - սկզբնաժամկետ'
'EndDate - վերջնաժամկետ
'CLimit - փոփոխված սահմանաչափ

Function Check_Changed_Limit(StartDate,EndDate,CLimit)
  Dim isEquale
  
  BuiltIn.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_ViewEdit & "|" & c_Other & "|" & c_Limits)
  BuiltIn.Delay(2000)
  'Կրացնում է սկզբնաժամկետ դաշտը
  Call Rekvizit_Fill("Dialog",1,"General","START","![End]" & "[Del]" & StartDate )
  'Լրացնում է վերջնաժամկետ դաշտը
  Call Rekvizit_Fill("Dialog",1,"General","END",EndDate )
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  wMDIClient.Refresh
  'Անցում է կատարում հաջորդ տողին
  wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").MoveLast
  'Կատարում է ստուգում ,եթե հավասար են վերադաձնում է TRUE ,հակառակ դեպքում դուրս է բերում սխալ
  isEquale = False
  If Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(3).Text) = Trim(CLimit) Then
      isEquale = True
  Else
      Log.Error("The Limitt doesn't change")
  End If
  Check_Changed_Limit =  isEquale 
  
End Function 
 
'_________________________________________________________________________________________________________________________________________________________
'Ջնջում է բոլոր ստեղծված փաստաթղթերը
'_________________________________________________________________________________________________________________________________________________________
'StartDate - սկզբնաժամկետ
'EndDate - վեջնաժամկետ
'Param - թղթապանակի անուն

Sub Delete_Actions(StartDate,EndDate,actionExists,actionType,Param)
 
   BuiltIn.Delay(delay_middle)
   wMainForm.Refresh
   BuiltIn.Delay(3000)
   Call wMainForm.MainMenu.Click(c_AllActions)
   Call wMainForm.PopupMenu.Click(Param)
   'Լրացնում է սկզբնաժամկետ դաշտը
   Call Rekvizit_Fill("Dialog",1,"General","START","![End]" & "[Del]" & StartDate)
   'Լրացնում է վերջնաժամկետ դաշտը
   Call Rekvizit_Fill("Dialog",1,"General","END","![End]" & "[Del]" & EndDate)
   '՚Լրացնում է Գործողության տեսակ դաշտը  
   If actionExists Then
     Call Rekvizit_Fill("Dialog",1,"General","DEALTYPE",  "^A[Del]" & actionType)
   End If
   Call ClickCmdButton(2, "Î³ï³ñ»É")
  
   wMDIClient.Refresh
   BuiltIn.Delay(4000)
   
   'Անցնում է ցուցակի մեջով `վեջից սկսած , և ջնջում է ամբողջը
   'Անցում է կատարում վերջին տողին
   wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").MoveLast
   Do Until wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").ApproxCount = 0
       'Կատարում է ջնջել գործողությունը
       BuiltIn.Delay(5000)
       Call wMainForm.MainMenu.Click(c_AllActions)
       Call wMainForm.PopupMenu.Click(c_Delete)
       If p1.WaitVBObject("frmAsMsgBox", 2500).Exists Then
          p1.VBObject("frmAsMsgBox").VBObject("cmdButton").Click()
          Call ClickCmdButton(3, "²Ûá")
          BuiltIn.Delay(7000)
       Else  
          Call ClickCmdButton(3, "²Ûá")
          If p1.WaitVBObject("frmAsMsgBox", 2500).Exists Then
             p1.VBObject("frmAsMsgBox").VBObject("cmdButton").Click()
             BuiltIn.Delay(3000)
             wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Refresh
             wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").MovePrevious
          End If
       End If       
    Loop
     'Փակում է պատուհանը
     BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_Windows)
    Call wMainForm.PopupMenu.Click(c_ClCurrWindow)
    BuiltIn.Delay(2000)
    
End Sub 
 
'_________________________________________________________________________________________________________________________________________________________
'Ջնջում է բոլոր ստեղծված փաստաթղթերը տրված քանակով օր՜ count = 0 ապա ջնջում է բոլորը
'_________________________________________________________________________________________________________________________________________________________
'StartDate - սկզբնաժամկետ
'EndDate - վեջնաժամկետ
'Param - թղթապանակի անուն

Sub Delete_Actions_ByCount(StartDate,EndDate,actionExists,actionType,Param,Count)
 
   Dim wMainForm,wMDIClient
   Set wMainForm = Sys.Process("Asbank").vbObject("MainForm")
   Set wMDIClient = wMainForm.Window("MDIClient")
   BuiltIn.Delay(delay_middle)
   wMainForm.Refresh
   BuiltIn.Delay(3000)
   Call wMainForm.MainMenu.Click(c_AllActions)
   Call wMainForm.PopupMenu.Click(Param)
   'Լրացնում է սկզբնաժամկետ դաշտը
   Call Rekvizit_Fill("Dialog",1,"General","START","![End]" & "[Del]" & StartDate)
   'Լրացնում է վերջնաժամկետ դաշտը
   Call Rekvizit_Fill("Dialog",1,"General","END","![End]" & "[Del]" & EndDate)
   '՚Լրացնում է Գործողության տեսակ դաշտը  
   If actionExists Then
     Call Rekvizit_Fill("Dialog",1,"General","DEALTYPE",actionType)
   End If
   Call ClickCmdButton(2, "Î³ï³ñ»É")
  
   Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).Refresh
   BuiltIn.Delay(2000)
   
   'Անցնում է ցուցակի մեջով `վեջից սկսած , և ջնջում է ամբողջը
   'Անցում է կատարում վերջին տողին
   Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient").VBObject("frmPttel_2").VBObject("tdbgView").MoveLast
   Do Until wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").ApproxCount = Count
             'Կատարում է ջնջել գործողությունը
             BuiltIn.Delay(2000)
             Call wMainForm.MainMenu.Click(c_AllActions)
             Call wMainForm.PopupMenu.Click(c_Delete)
             If Sys.Process("Asbank").WaitVBObject("frmAsMsgBox", 1500).Exists Then
                Sys.Process("Asbank").VBObject("frmAsMsgBox").VBObject("cmdButton").Click()
                Call ClickCmdButton(3, "²Ûá")
                BuiltIn.Delay(7000)
             Else  
                Call ClickCmdButton(3, "²Ûá")
                If Sys.Process("Asbank").WaitVBObject("frmAsMsgBox", 1500).Exists Then
                   Sys.Process("Asbank").VBObject("frmAsMsgBox").VBObject("cmdButton").Click()
                   wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").MovePrevious
                End If
             End If       
    Loop
     'Փակում է պատուհանը
     BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_Windows)
    Call wMainForm.PopupMenu.Click(c_ClCurrWindow)
    
End Sub 
'______________________________________________________________________________________
'Ջնջում է գլխավոր պայմանագիրը
'______________________________________________________________________________________

Sub Delete_Doc()
   
   Dim wMainForm,wMDIClient
   Set wMainForm = Sys.Process("Asbank").vbObject("MainForm")
   Set wMDIClient = wMainForm.Window("MDIClient")
   BuiltIn.Delay(delay_middle)
   wMainForm.Refresh
   Call wMainForm.MainMenu.Click(c_AllActions)
   Call wMainForm.PopupMenu.Click(c_Delete)
   Call ClickCmdButton(3, "²Ûá")
   
End Sub 


'----------------------------------------------------------------------------------------------
' Կատարում է action-ով տրված գործողությունը
'----------------------------------------------------------------------------------------------
Sub ContractAction (action)
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(action)
End Sub

'---------------------------------------------------------------------------------------------
' Փնտրում է name անունով փաստաթուղթը ըստ i-րդ սյան արժեքի։
' Վերադարձնում է True, եթե այն գտնվել է, False` հակառակ դեպքում։
'---------------------------------------------------------------------------------------------
Function Find_Data(name, i)
  Dim Exists
  Exists = False
  Set my_vbObj = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).WaitVBObject("frmPttel", 500000)
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).Refresh

  BuiltIn.delay(2000)
  If my_vbObj.Exists Then 
  my_vbObj.vbObject("tdbgView").MoveFirst
     Do While (Not my_vbObj.vbObject("tdbgView").EOF)
      If Trim(my_vbObj.vbObject("tdbgView").Columns.Item(i).Text) = Trim(name) then 
        Exists = True 
        Exit Do   
      Else
        Call my_vbObj.vbObject("tdbgView").MoveNext
      End If
    Loop 
  Else
    Log.Error("Թղթապանակը հնարավոր չեղավ բացել")
  End If

  Find_Data = Exists
        
End Function

'---------------------------------------------------------------------------------------------
' Վարկային գծի դադարեցում/ վերականգնում փաստաթղթի լրացում
' TerRes - Վարկային գծի դադարեցում/ վերականգնում
' Date - Ամսաթիվ
'Val - Գծայնության կարգավիճակ (1,2)
'---------------------------------------------------------------------------------------------
Sub Credit_Termination_Restoration(TerRes, Date, Val)
  'Կատարել "Գործողություններ/Բոլոր գործողություններ/Պայմաններ և վիճակներ/Գծայնության վերականգնում(դադարեցում)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_TermsStates & TerRes)
  wMDIClient.Refresh
  
  'Լրացնում է ամսաթիվ դաշտը
  Call Rekvizit_Fill("Document",1,"General","DATE","![End]" & "[Del]" & Date)
  'Լրացնում է գծայնության կարգավիճակ դաշտը
  Call Rekvizit_Fill("Document",1,"General","LNBR",Val)
  Call ClickCmdButton(1, "Î³ï³ñ»É")
End Sub

'---------------------------------------------------------------------------------------------
' Ակցեպտավորում փաստաթղթի ստեղծում
' Date --- Ամսաթիվ
' Money --- Գումար
' ReqFDate --- Պահանջի ժամկետ
' OblFDate --- Պարտավորության ժամկետ
' OblPer --- Պարտավուրության տոկոսադրույք
' Baj --- բաժ.
' Bank --- Ծանուցող բանկ 
' Acc --- Հաշվարկային հաշիվ
'---------------------------------------------------------------------------------------------
Sub Create_Acceptance(Date, Money, ReqFDate, OblFDate, OblPer, Baj, Bank, Acc)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Acceptance) 
  
  wMDIClient.Refresh 
  
  'Լրացնել "Ամսաթիվ " դաշտը Date արժեքով
  Call Rekvizit_Fill("Document",1,"General","DATE","![End]" & "[Del]" & Date)
  'Լրացնել "Գումար" դաշտը Money արժեքով
  Call Rekvizit_Fill("Document",1,"General","SUMMA",Money) 
  'Լրացնել "Պահանջի ժամկետ " դաշտը ReqFDate արժեքով
  Call Rekvizit_Fill("Document",1,"General","CDATEAGR",ReqFDate)
  'Լրացնել "Պարտավորության ժամկետ " դաշտը OblFDate արժեքով
  Call Rekvizit_Fill("Document",1,"General","DDATEAGR",OblFDate)
  'Լրացնել "Պարտավուրության տոկոսադրույք" դաշտը OblPer արժեքով 
  Call Rekvizit_Fill("Document",1,"General","DPCAGR",OblPer)
  'Լրացնել "բաժ." դաշտը Baj արժեքով
  Call Rekvizit_Fill("Document",1,"General","DPCAGR",Baj)
  'Լրացնել "Ծանուցող բանկ" դաշտը Bank արժեքով
  Call Rekvizit_Fill("Document",1,"General","CLIBANK",Bank)
  'Լրացնել "Հաշվարկային հաշիվ " դաշտը Acc արեքով
  Call Rekvizit_Fill("Document",1,"General","ACCACC",Acc)
  'Սեղմել "Կատարել " կոճակը 
  Call ClickCmdButton(1, "Î³ï³ñ»É")
End Sub

'---------------------------------------------------------------------------------------------
' Վերադարձնում է ակցեպտավորում փաստաթղթի ColNum համարով սյան արժեքը
'ColNum --- սյունի համար
'---------------------------------------------------------------------------------------------
Function FindVal (ColNum)
  'Կատարել "Գործողություններ/ Բոլոր գործողություններ/ Թղթապանակներ/ Պարտավորություններ ակրեդիտիվի գծով" գործողությունը
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_LiabilitiesInLC)
   
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).Refresh
  Set my_vbObj = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).WaitVBObject("frmPttel_2", 50000)
  
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).Refresh
  If my_vbObj.Exists Then 
   FindVal = Trim(my_vbObj.vbObject("tdbgView").Columns.Item(ColNum).Text)
  End If
  
  'փաստաթղթի համարը վերցնելուց հետո փակենլ  Պարտավորություններ ակրեդիտիվի գծով պատուհանը
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_2").Close
  'փակենլ Պայմանագրեր/Ակրեդիտիվ/ պատուհանը
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
End Function 

'---------------------------------------------------------------------------------------------
' Ակրեդիտիվ պայմանագրի պարտքերի մարման հայտի ստեղծում
' Date --- Ամսաթիվ
' Money --- Հիմնական գումար
' Count --- Հաշիվ 
' AMDCount --- Դրամային հաշիվ
'fBase - պարտքերի մարման 
'-----------------------------------------------------------------
Sub AkrPayment(fBase , Date, Money, Count, AMDCount)
  'Կատարել "Գործողություններ/Բոլոր Գործողություններ/Գործողություններ/Տրամադրում Մարում/Պարտքերի մարում"
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_PayOffDebt)
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).Refresh
  
  fBASE = Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmASDocForm").DocFormCommon.Doc.isn
  
  'Լրացնել Ամսաթիվ դաշտը Date արժեքով
  Call Rekvizit_Fill("Document",1,"General","DATE","![End]" & "[Del]" & Date)
  'Լրացնել Հիմնական գումար դաշտը Money  արժեքով
  Call Rekvizit_Fill("Document",1,"General","SUMAGR",Money)
  'Լրացնել Հաշիվ դաշտը Count արժեքով
  Call Rekvizit_Fill("Document",1,"General","ACCCORR",Count)
  'Լրացնել Դրամային հաշիվ դաշտը AMDCount արժեքով
  Call Rekvizit_Fill("Document",1,"General","AMDACCCORR",AMDCount)
  'Սեղմել "Կատարել " կոճակը - Հաշվի մնացորդը պետք է զրոյանա :
  Call ClickCmdButton(1, "Î³ï³ñ»É")
  Call ClickCmdButton(5, "²Ûá")
End Sub

'---------------------------------------------------------------------------------------------
' Ջնջում է  փաստաթղթերը Count-ին համապատասխան
'---------------------------------------------------------------------------------------------
Sub Delete_ByCount(Count)
  'Կատարել "Գործողություններ/Բոլոր Գործողություններ/Թղթապանակներ/ Ակցեպտավորումներ/"
  'Կատարել "Գործողություններ/Բոլոր Գործողություններ/Գործողությունների դիտում"
  'Կատարել "Գործողություններ/Ջնջել"
      
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_Acceptances)
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_OpersView)
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate").keys("^A[Del]" & "[Tab]")
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate_2").Keys("^A[Del]" & "[Tab]")
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
    
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).Refresh
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_3").VBObject("tdbgView").MoveLast    
    Do While Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_3").VBObject("tdbgView").VisibleRows<>Count  
      Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
      Sys.Process("Asbank").VBObject("frmDeleteDoc").VBObject("YesButton").ClickButton
    Loop
 
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_3").Close   
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_2").Close
   
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ViewEdit & "|" & c_Other & "|" & c_CalcDates)
    'ջնջել <<Հաշվարկման ամսաթվեր>>-ից
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate").keys("^A[Del]" & "[Tab]")
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate_2").Keys("^A[Del]" & "[Tab]")
    Call ClickCmdButton(2, "Î³ï³ñ»É")

    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).Refresh
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_2").VBObject("tdbgView").MoveLast    
    Do While Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_2").VBObject("tdbgView").VisibleRows<>Count  
      Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
      Call ClickCmdButton(3, "²Ûá")
    Loop
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_2").Close
  
    'ջնջել <<ê³ÑÙ³Ý³ã³÷»ñ>>-ից   
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ViewEdit & "|" & c_Other & "|" & c_Limits)
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate").keys("^A[Del]" & "[Tab]")
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate_2").Keys("^A[Del]" & "[Tab]")
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("Checkbox").Value = 1
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
  
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).Refresh
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_2").VBObject("tdbgView").MoveLast    
    While (Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_2").VBObject("tdbgView").VisibleRows<>Count) 
    Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
      Call ClickCmdButton(3, "²Ûá")
    Wend
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_2").Close
  
End Sub

'-------------------------------------------------------------------------------------
' Բարդ վարկային գծի համար Պահուստավորման գործողության կատարում
'--------------------------------------------------------------------------------------
' DocN - Պայմանագրի համարը (V-002457)
' res_unres - 1 արծեքի դեպքում կատարում է պահուստավորում , 0-ի դեպքում ապապահուստավորում
' res_date - Պահուստավորման/Ապապահուստավորման գործողության ամսաթիվ
' res_sum - Պահուստավորման/Ապապահուստավորման գործողության գումար
Sub Compl_Actions_Reservation(res_unres, res_date, res_sum)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_Store & "|" & c_UnusedPartStore)
    
    Call Rekvizit_Fill("Document", 1, "General", "DATE", res_date) 

    If res_unres = 1 Then
      Call Rekvizit_Fill("Document", 1, "General", "SUMRNC", res_sum) 
    Else
      Call Rekvizit_Fill("Document", 1, "General", "SUMUNRNC", res_sum) 
    End If

    Call ClickCmdButton(1, "Î³ï³ñ»É")
    Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmPttel").Close
End Sub