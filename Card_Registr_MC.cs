Option Explicit
'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_Common
'USEUNIT Card_Library
'USEUNIT Mortgage_Library
'USEUNIT OLAP_Library
'USEUNIT Constants
'USEUNIT Library_Colour

'Test Case Id 165987

'Պլաստիկ Քարտեր ԱՇՏ/MC Ստացված գործողություններ
Sub Cards_Registr_MC_Test()
    
    Dim DateStart,DateEnd, filesSource, FSO
    Dim MCReceivedTrans,Path1,Path2,resultWorksheet
    Dim queryString,sql_Value, colNum,sql_isEqual,result,expValues
  
    DateStart = "20010101"
    DateEnd = "20240101"

    Set MCReceivedTrans = New_MCReceivedTrans()
    With MCReceivedTrans
        .FileDate_1 = "121017"
        .FileDate_2 = "121017"
        .ShowAllTransactions = 1
        .ShowArchivedOpers = 0
    End With
    queryString = "update statistics HI DELETE FROM HI WHERE fDATE = '2018-05-24'"
  
    Call Initialize_AsBankQA(DateStart, DateEnd) 

    Call Create_Connection()
    Call Execute_SLQ_Query(queryString)
    filesSource = Project.Path & "Stores\Import_Files\"
    Call SetParameter("SVRECEIVE", filesSource)
    Login("ARMSOFT")
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CopyFile Project.Path & "Stores\Import_Files_For_Copy\*", Project.Path & "Stores\Import_Files\", True
    Set FSO = Nothing
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''--- Ֆայլերի ընդունում ---'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    Log.Message "--- Add Files ---" ,,, DivideColor  
    
    Call ChangeWorkspace(c_CardsSV)
    
    Call wTreeView.DblClickItem("|äÉ³ëïÇÏ ù³ñï»ñÇ ²Þî (SV)|ü³ÛÉ»ñÇ ÁÝ¹áõÝáõÙ")
    Call ClickCmdButton(5, "²Ûá")
    BuiltIn.Delay(8000)  
    wMDIClient.VBObject("FrmSpr").Close()
  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''--- Թղթապանակի ստուգում ֆայլերի ընդունելուց հետո ---''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' ''''''''''''''
    Log.Message "--- Check MCReceived Transaction After Add Files ---" ,,, DivideColor 

    Call GoToMCReceivedTrans_PlasticCarts(MCReceivedTrans) 
    Call CheckPttel_RowCount("frmPttel", 214)
    Call wMainForm.MainMenu.Click(c_Views & "|" & "Սորտավորած BO UTRNNO")
    BuiltIn.Delay(2000)
    
    Path1 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Actual\Actual_MCReceived.xlsx"
    Path2 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Expected\Expected_MCReceived.xlsx"
    resultWorksheet = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Result\Result_MCReceived.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''--- Հաշվառել փաստաթղթերը ---''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    Log.Message "--- Registr Cards Files ---" ,,, DivideColor 

    Call Registr_Cards_Total("240518")
    Call Close_Pttel("frmPttel")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''--- Թղթապանակի ստուգում ֆայլերի Հաշվառելուց հետո ---'''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''' ''''''''''''''
    Log.Message "--- Check MCReceived Transaction After Registr Cards ---" ,,, DivideColor 

    MCReceivedTrans.FileDate_2 = "240518"
    
    Call GoToMCReceivedTrans_PlasticCarts(MCReceivedTrans) 
    Call CheckPttel_RowCount("frmPttel", 214)
    Call wMainForm.MainMenu.Click(c_Views & "|" & "Սորտավորած BO UTRNNO")
    BuiltIn.Delay(2000)
    
    Path1 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Actual\Actual_MCReceived_AfterRegistr.xlsx"
    Path2 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Expected\Expected_MCReceived_AfterRegistr.xlsx"
    resultWorksheet = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Result\Result_MCReceived_AfterRegistr.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''--- Կատարում է SQL ստուգում ---''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    Log.Message "SQL Check For MCReceived Transaction",,,SqlDivideColor
    
      'Կատարում է SQL ստուգում
      queryString = "select Count(*) from HI where fDATE = '2018-05-24' "
      sql_Value = 23
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
   
      queryString = "select Sum(fSUM) from HI where fDATE = '2018-05-24' and fTYPE = '01' and fDBCR = 'C' "
      sql_Value = 58851.07
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
  
      queryString = "select Sum(fCURSUM) from HI where fDATE = '2018-05-24' and fTYPE = '01' and fDBCR = 'C'"
      sql_Value = 58851.07
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
       
      queryString = "select Sum(fSUM) from HI where fDATE = '2018-05-24' and fTYPE = '01' and fDBCR = 'D'"
      sql_Value = 58851.07
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If 
       
      queryString = "select Sum(fCURSUM) from HI where fDATE = '2018-05-24' and fTYPE = '01' and fDBCR = 'D'"
      sql_Value = 38220.3
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If      
       
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''--- Ջնջում է բոլոր ներմուծած ֆայլերը ---''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Delete All Contracts Total ---" ,,, DivideColor       

    'Ջնջում է բոլոր ներմուծած ֆայլերը
    Call Delete_All_Contracts_Total()
    Call Close_Pttel("frmPttel")
  
      'Կատարում է SQL ստուգում
      queryString = "select Count(*) from HI where fDATE = '2018-05-24' "
      sql_Value = 0
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
       
    'Փակել ASBANK-ը
    Call Close_AsBank()
End Sub