Option Explicit
'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Online_PaySys_Library
'USEUNIT CashInput_Confirmphases_Library
'USEUNIT Currency_Exchange_Confirmphases_Library
'USEUNIT BackBallance_Input_Confirmphases_Library

'Test case number - 17094

Sub BackBallance_Input_Allconditions_Test()
    
    Dim fDATE, startDATE , docNumber, summa, fISN, draft,aim, nbAcc, data
    Dim confInput, confPath, docExist, isDel, rCount
    data = null
    Utilities.ShortDateFormat = "yyyymmdd"
    startDATE = "20120101"
    fDATE = "20130101"
    confPath = "X:\Testing\BackBallance_Input confirm phases\BackBallance_Input_Allconditions.txt"
    
    
    nbAcc = "999998/900002"
    summa = "250000"
    draft = False
    aim = "Black list"
    
    'Test StartUp start
    Call Initialize_AsBank("bank", startDATE, fDATE)
    'Î³ñ·³íáñáõÙÝ»ñÇ Ý»ñÙáõÍáõÙ
    confInput = Input_Config(confPath)
    If Not confInput Then
        Log.Error("The configuration doesn't input")
    End If
    'Test StartUp end
    
     Call ChangeWorkspace("Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ (ÁÝ¹É³ÛÝí³Í)")
     
    Call Online_PaySys_Go_To_Agr_WorkPapers("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ", data, data)
    'ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ·ñÇ ëï»ÕÍáõÙ
    Call CashInput_Doc_Fill(docNumber, nbAcc, summa, aim, fISN, draft)                
    
    Sys.Process("Asbank").vbObject("frmAsMsgBox").vbObject("cmdButton").Click()
    'îå»Éáõ Ó¨ å³ïáõÑ³ÝÇ ÷³ÏáõÙ
    BuiltIn.Delay(delay_small)
    Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("FrmSpr").Close
    Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmPttel").Close
    
    'ö³ëï³ÃÕÃÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý ë¨ óáõó³Ï
    Call Online_PaySys_Send_To_Verify(3)
    
    'ê¨ óáõó³ÏÇó ÷³ëï³ÃÕÃÇ Ñ³ëï³ïáõÙ Ñ³ÙÁÝÏÝáõÙÝ»ñÇ Ù³ëÇÝ ÇÝýáñÙ³óÇ³Ý ëïáõ·»Éáõó Ñ»ïá
    
    Call ChangeWorkspace("§ê¨ óáõó³Ï¦ Ñ³ëï³ïáÕÇ ²Þî")
    docExist = Online_PaySys_Check_Doc_In_Black_List(docNumber)
    If docExist = False Then
        Log.Error("Document with specified ID " & docNumber & "doesn't exists in Black list folder")
        Exit Sub
    End If
    rCount = Online_PaySys_Check_Assertion_In_Black_List()
    If rCount <> 2 Then
        Log.Error("There must be 2 row")
        Exit Sub
    End If
    
    Call Online_PaySys_Verify( True)
    
    
    'ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ
    Login("VERIFIER")
    'ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ 1-ÇÝ Ñ³ëï³ïáÕÇ Ùáï
    docExist = Online_PaySys_Check_Doc_In_Verifier(docNumber, null, Null)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " doesn't exist in 1st verify documents")
        Exit Sub
    End If
    'ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ 1-ÇÝ Ñ³ëï³áïÕÇ ÏáÕÙÇó
    Call Online_PaySys_Verify(True)
    
    'ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ ²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ ÃÕÃ³å³Ý³ÏÇó
    Login("ARMSOFT")
    Call ChangeWorkspace("Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ (ÁÝ¹É³ÛÝí³Í)")
    'ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ ÃÕÃ³å³Ý³ÏáõÙ
    docExist = Online_PaySys_Check_Doc_In_Workpapers(docNumber, null, Null)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " doesn't exist in workpaper documents")
        Exit Sub
    End If
    Call PaySys_Verify(True)
    Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmPttel").Close
    
    '¶ÉË³íáñ Ñ³ßí³å³ÑÇ ÁÝ¹Ñ³Ýáõñ ¹ÇïáõÙ ÃÕÃ³å³Ý³ÏáõÙ ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ
    Call ChangeWorkspace("¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî")
    docExist = Check_Doc_In_GeneralView_Folder(fISN)
    If Not docExist Then
        Log.Error("The document with number " & fISN & " must exist in general view folder")
        Exit Sub
    End If
    
    'Test CleanUp start
    'ö³ëï³ÃÕÃÇ Ñ»é³óáõÙ
    Call Online_PaySys_Delete_Agr()
    Call Close_AsBank()
    'Test CleanUp End
End Sub