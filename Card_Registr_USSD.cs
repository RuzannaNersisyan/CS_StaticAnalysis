Option Explicit
'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_Common
'USEUNIT Card_Library
'USEUNIT Mortgage_Library
'USEUNIT OLAP_Library
'USEUNIT Constants
'USEUNIT Library_Colour

'Test Case Id 165999

'Պլաստիկ Քարտեր ԱՇՏ/USSD Հաշվետվություն թղթապանակ
Sub Cards_Registr_USSD_Test()

    Dim DateStart,DateEnd
    Dim USSD_Reports,Path1,Path2,resultWorksheet,SortArr(0)
    Dim queryString,sql_Value, colNum,sql_isEqual
  
    DateStart = "20010101"
    DateEnd = "20240101"

    Set USSD_Reports = New_USSD_Reports()
    With USSD_Reports
        .FileDate_1 = "250317"
        .FileDate_2 = "250317"
        .ShowProcessed = 1
        .View = "vACUSSD\2"
        .FillInto = "0"
    End With

    queryString = "update statistics HI DELETE FROM HI WHERE fDATE = '2018-05-22'"
  
    'Test StartUp start
    Call Initialize_AsBankQA(DateStart, DateEnd) 

    Call Create_Connection()
    Call Execute_SLQ_Query(queryString)
    Call ChangeWorkspace(c_CardsSV)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''--- Թղթապանակի ստուգում ֆայլերի ընդունելուց հետո ---''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' ''''''''''''''
    Log.Message "--- Check USSD Reports After Add Files ---" ,,, DivideColor 

    Call GoToUSSD_Reports_PlasticCarts(USSD_Reports)
    Call CheckPttel_RowCount("frmPttel", 5)

    Path1 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Actual\Actual_USSD_Reports.xlsx"
    Path2 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Expected\Expected_USSD_Reports.xlsx"
    resultWorksheet = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Result\Result_USSD_Reports.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''--- Հաշվառել փաստաթղթերը ---''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    Log.Message "--- Registr Cards Files ---" ,,, DivideColor 

    Call Registr_Cards_Total("260518")
    Call Close_Pttel("frmPttel")   
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''--- Թղթապանակի ստուգում ֆայլերի Հաշվառելուց հետո ---'''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''' ''''''''''''''
    Log.Message "--- Check USSD Reports After Registr Cards ---" ,,, DivideColor 

    USSD_Reports.FileDate_2 = "260518"
    Call GoToUSSD_Reports_PlasticCarts(USSD_Reports)
    Call CheckPttel_RowCount("frmPttel", 5)

    Path1 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Actual\Actual_USSD_ReportsAfterRegistr.xlsx"
    Path2 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Expected\Expected_USSD_ReportsAfterRegistr.xlsx"
    resultWorksheet = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Result\Result_USSD_ReportsAfterRegistr.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()    
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''--- Կատարում է SQL ստուգում ---''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    Log.Message "SQL Check For USSD Reports",,,SqlDivideColor    
         
        'Կատարում է SQL ստուգում
        queryString = "select Count(*) from HI where fDATE = '2018-05-26' "
        sql_Value = 6
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select Sum(fSUM) from HI where fDATE = '2018-05-26' and fTYPE = '01' and fDBCR = 'C' "
        sql_Value = 250
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select Sum(fCURSUM) from HI where fDATE = '2018-05-26' and fTYPE = '01' and fDBCR = 'C'"
        sql_Value = 250
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
         
        queryString = "select Sum(fSUM) from HI where fDATE = '2018-05-26' and fTYPE = '01' and fDBCR = 'D'"
        sql_Value = 250
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If 
         
        queryString = "select Sum(fCURSUM) from HI where fDATE = '2018-05-26' and fTYPE = '01' and fDBCR = 'D'"
        sql_Value = 250
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
      queryString = "select Count(*) from HI where fDATE = '2018-05-26' "
      sql_Value = 0
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
       
    'Փակել ASBANK-ը
    Call Close_AsBank()
End Sub