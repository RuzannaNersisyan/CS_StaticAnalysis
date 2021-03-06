Option Explicit

'USEUNIT Library_Common
'USEUNIT OLAP_Library
'USEUNIT Constants

'Test Case ID 166359 

Sub Accountanty_31_Test() 
    
    Dim exists,SPath,userName ,Passord ,StartDate,EndDate,AccNumber ,TreeLevel,CBranch,Language 
    Dim EPath1,EPath2,Thousand,RequesQuery, i ,j,DB1,param,SheetN ,CurrDate,windExists,Cont
    Dim resultWorksheet(4)
    Dim DateStart, DateEnd, eDate, sDate
     
    SPath = Project.Path & "Stores\Actual_OLAP"
    EPath1 = Project.Path & "Stores\Actual_OLAP\16600_31.xls"
    EPath2 = Project.Path & "Stores\Expected_OLAP\16600_31_01.02.2014-28.02.2014_cragri nersic_ok.xls"
    For i = 1 To 4
      resultWorksheet(i) = Project.Path & "Stores\Result_Olap\Result_16600_31_sheet_" & i  & ".xls"
    Next
   'Î³ï³ñáõÙ ¿ ëïáõ·áõÙ,»Ã» ÝÙ³Ý ³ÝáõÝáí ý³ÛÉ Ï³ ïñí³Í ÃÕÃ³å³Ý³ÏáõÙ ,çÝçáõÙ ¿   
    exists = aqFile.Exists(EPath1)
    If exists Then
        aqFileSystem.DeleteFile(EPath1)
    End If

    DateStart = "20120101"
    DateEnd = "20240101"
    param = "16600_31.xls"
   
    Call Initialize_AsBankQA(DateStart, DateEnd) 
 
    Call ChangeWorkspace(c_ChiefAcc)
    Call wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|Î´ Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|31 îñ³Ù³¹ñí³Í ù³ñï»ñÇ,ù³ñï³ÛÇÝ ·áñÍ³éÝáõÃÛáõÝÝ»ñÇ ¨ ëå³ë.ÙÇçáóÝ»ñÇ í»ñ³µ»ñÛ³É")
    sDate = "010214"
    eDate = "280214"
    Cont = False
    BuiltIn.Delay(3000)
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate").Keys(sDate & "[Tab]")
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate_2").Keys(eDate & "[Tab]")
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").Click()
   
    BuiltIn.Delay(10000)
    'ä³Ñ»É ý³ÛÉÁ ACTUAL_OLAP ÃÕÃ³å³Ý³ÏáõÙ
    Call Save_To_Folder(SPath,param,Cont)
    'Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call  CompareTwoExcelFiles(EPath1, EPath2, resultWorksheet)
    
    'Î³ï³ñ»É ²ßË³ï³ÝùÇ ³í³ñï
    Sys.Process("EXCEL").Window("XLMAIN", "" & param & "  [Compatibility Mode] - Excel", 1).Window("EXCEL2", "", 2).ToolBar("Ribbon").Window("MsoWorkPane", "Ribbon", 1).Window("NUIPane", "", 1).Window("NetUIHWND", "", 1).Keys("~X")
    Sys.Process("EXCEL").Window("XLMAIN", "" & param & "  [Compatibility Mode] - Excel", 1).Window("EXCEL2", "", 2).ToolBar("Ribbon").Window("MsoWorkPane", "Ribbon", 1).Window("NUIPane", "", 1).Window("NetUIHWND", "", 1).Keys("Y7")
 
    'ö³Ï»É EXCEL- Á
    Call CloseAllExcelFiles()
    Call Close_AsBank()
'    TestedApps.killproc.Run()
  
End Sub