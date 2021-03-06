Option Explicit
'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_Common
'USEUNIT Card_Library
'USEUNIT Mortgage_Library
'USEUNIT OLAP_Library
'USEUNIT Constants
'USEUNIT Library_Colour

'Test Case ID 165989

'Պլաստիկ Քարտեր ԱՇՏ/Համատեղ տերմինալ գործողություններ թղթապանակ
Sub Cards_Registr_Terminal_Test()

    Dim DateStart,DateEnd
    Dim SharedTerminalOperation,Path1,Path2,resultWorksheet
    Dim queryString,sql_Value, colNum,sql_isEqual,result,expValues
    
    DateStart = "20010101"
    DateEnd = "20240101"
    
    Set SharedTerminalOperation = New_SharedTerminalOperations()
    With SharedTerminalOperation
        .FileDate_1 = "121017"
        .FileDate_2 = "121017"
        .ShowMadeTransactions = 1
        .ShowArchivedOpers = 0
        .ShowInsideBankingCommission = 1
        .View = "VrtTrns\2"
        .FillInto = "0"
    End With
    Dim SortArr(1)
    SortArr(0) = "FILEDATE"

    queryString = "update statistics HI DELETE FROM HI WHERE fDATE = '2018-05-25'"
  
    'Test StartUp start
    Call Initialize_AsBankQA(DateStart, DateEnd) 

    Call Create_Connection()
    Call Execute_SLQ_Query(queryString)
    Call ChangeWorkspace(c_CardsSV)
  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''--- Թղթապանակի ստուգում ֆայլերի ընդունելուց հետո ---''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' ''''''''''''''
    Log.Message "--- Check Shared Terminal Operation After Add Files ---" ,,, DivideColor 

    Call GoToSharedTermOperation_PlasticCarts(SharedTerminalOperation) 
    Call CheckPttel_RowCount("frmPttel", 50)
    
    Path1 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Actual\Actual_SharedTerminalOperation.xlsx"
    Path2 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Expected\Expected_SharedTerminalOperation.xlsx"
    resultWorksheet = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Result\Result_SharedTerminalOperation.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''--- Հաշվառել փաստաթղթերը ---''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    Log.Message "--- Registr Cards Files ---" ,,, DivideColor 

    Call Registr_Cards_Total("220518")
    Call Close_Pttel("frmPttel")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''--- Թղթապանակի ստուգում ֆայլերի Հաշվառելուց հետո ---'''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''' ''''''''''''''
    Log.Message "--- Check Shared Terminal Operation After Registr Cards ---" ,,, DivideColor 

    SharedTerminalOperation.FileDate_2 = "220518"
    
    Call GoToSharedTermOperation_PlasticCarts(SharedTerminalOperation)  
    Call CheckPttel_RowCount("frmPttel", 50)

    Path1 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Actual\Actual_SharedTerminalOperation_AfterRegistr.xlsx"
    Path2 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Expected\Expected_SharedTerminalOperation_AfterRegistr.xlsx"
    resultWorksheet = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Result\Result_SharedTerminalOperation_AfterRegistr.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''--- Կատարում է SQL ստուգում ---''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    Log.Message "SQL Check For Shared Terminal Operation Transaction",,,SqlDivideColor    
    
      'Կատարում է SQL ստուգում
      queryString = "select Count(*) from HI where fDATE = '2018-05-22' "
      sql_Value = 8
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
       
      queryString = "select Sum(fSUM) from HI where fDATE = '2018-05-22' and fTYPE = '01' and fDBCR = 'C' "
      sql_Value = 62000.00
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
  
      queryString = "select Sum(fCURSUM) from HI where fDATE = '2018-05-22' and fTYPE = '01' and fDBCR = 'C'"
      sql_Value = 62000.00
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
       
      queryString = "select Sum(fSUM) from HI where fDATE = '2018-05-22' and fTYPE = '01' and fDBCR = 'D'"
      sql_Value = 62000.00
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If 
       
      queryString = "select Sum(fCURSUM) from HI where fDATE = '2018-05-22' and fTYPE = '01' and fDBCR = 'D'"
      sql_Value = 62000.00
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
      queryString = "select Count(*) from HI where fDATE = '2018-05-25' "
      sql_Value = 0
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
       
    'Փակել ASBANK-ը
    Call Close_AsBank()
End Sub