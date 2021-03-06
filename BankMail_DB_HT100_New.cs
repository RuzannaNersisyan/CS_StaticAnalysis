Option Explicit
'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Online_PaySys_Library
'USEUNIT BankMail_Library
'USEUNIT BankMail_Library
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Library_Contracts

' Test Case ID 165070

Sub PayOrder_Verify_and_SendBankMail_DB_Test_New()
  
      Dim pattern, checkVal, wDate, lacsBranch, tdDate
      Dim ordType, fISN, wAcsBranch, wAcsDepart, payDate, docNum, cliCode, accDB, payer, ePayer, taxCods,_
              jurState, dbDropDown, coaNum, balAcc, accMask, accCur, accType, cliName, cCode, accNote, accNote2,_
              accNote3, acsBranch, acsDepart, acsType, pCardNum, socCard, accCR, receiver, eReceiver, summa, wCur,_
              wAim, jurStatR, bankCr, authorPerson, addInfo, wAddress, authPerson, rInfo
      Dim workEnvName, workEnv, stRekName, wDateS, endRekName, wDateE, wStatus, isnRekName
      Dim colN, action, doNum, doActio, status, state, ordChildISN
      Dim docTypeName, commentName, wDateTime
      Dim queryString, sqlValue, colNum, sql_isEqual
      Dim childISN, wChildISN, grRemOrdNum, grRemOrdISN, ordChildIS
      Dim paramName, paramValue, sBody, bodyValue, confInput, confPath
      Dim startDate, fDate, verifyDocuments
      
      startDate = "20030101"
      fDate = "20250101"
      Call Initialize_AsBank("bank", startDate, fDate)
                     
      ' Մուտք համակարգ ARMSOFT օգտագործողով
      BuiltIn.Delay(10000)
      Call Create_Connection()
      Login("ARMSOFT")
      
      BuiltIn.Delay(10000)
      paramName = "BMUSEDB"
      paramValue = "1"
      Call SetParameter(paramName, paramValue)
      
      ' BMDBSERVER պարամետրի արժեքը դնել qasql2017
      paramName = "BMDBSERVER"
      paramValue = "qasql2017"
      Call SetParameter(paramName, paramValue)
      
      ' BMDBNAME պարամետրի արժեքը դնել BankMail_Testing
      paramName = "BMDBNAME"
      paramValue = "BankMail_Testing"
      Call SetParameter(paramName, paramValue)
      
      ' Կարգավորումների ներմուծում
      confPath = "X:\Testing\Order confirm phases\BankMailNoVerify.txt"
      confInput = Input_Config(confPath)
      If Not confInput Then
          Log.Error("Կարգավորումները չեն ներմուծվել")
          Exit Sub
      End If
      
      ' Կարգավորումների ներմուծում
      confPath = "X:\Testing\Order confirm phases\AllVerify_New.txt"
      confInput = Input_Config(confPath)
      If Not confInput Then
          Log.Error("Կարգավորումները չեն ներմուծվել")
          Exit Sub
      End If

      ' Մուտք Հաճախորդների սպասարկում և դրամարկղ(Ընդլայնված)
      Call ChangeWorkspace(c_CustomerService)
      
      ' Մուտք աշխատանքային փաստաթղթեր
      workEnvName = "|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      workEnv = "Աշխատանքային փաստաթղթեր"
      wDate = "290517"
      stRekName = "PERN"
      endRekName = "PERK"
      wStatus = False
      state =  AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք Աշխատանքային փաստաթղթեր թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      payDate = "290517"
      ordType = "PAY"
      cliCode = "¶»Ý¹³Éý"
      accCR = "1810003359330100"
      receiver = "¶³É³¹ñÇ»É"
      summa = "390.00"
      wAim = "Ð³íÇïÛ³Ý (HT100 Ï³Ù HT203)"
      accDB = "7770000067030100"
      wCur = "000"
      dbDropDown = False
      payer = "ä»ïñáë äáÕáëÛ³Ý"
      taxCods = "12345678"
'      accMask = "00067030100"
      ' Վճարման հանձնարարագրի ստեղծում 
      Call PaymOrdToBeSentFill(ordType, fISN, wAcsBranch, wAcsDepart, payDate, docNum, cliCode, accDB, payer, ePayer, taxCods,_
                                                          jurState, dbDropDown, coaNum, balAcc, accMask, accCur, accType, cliName, cCode, accNote, accNote2,_
                                                          accNote3, acsBranch, acsDepart, acsType, pCardNum, socCard, accCR, receiver, eReceiver, summa, wCur,_
                                                          wAim, jurStatR, bankCr, authorPerson, addInfo, wAddress, authPerson, rInfo)
      Log.Message(fISN)
      Log.Message(docNum)
      Log.Message("SQL Check 1") 
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '11' AND fOBJECT = '1630509' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' " & _ 
                              " AND fCURSUM = '390.00' AND fOP = 'TRF' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '390.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '11' AND fOBJECT = '425737041' AND fCUR = '000' AND fDATE = '2017-05-29 00:00:00' " & _ 
                              " AND fCURSUM = '390.00' AND fOP = 'TRF' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '390.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '11' AND fOBJECT = '1630442' AND fCUR = '000' AND fDATE = '2017-05-29 00:00:00'  " & _ 
                              " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '11' AND fOBJECT = '425737041' AND fCUR = '000' AND fDATE = '2017-05-29 00:00:00' " & _ 
                              " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '500.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
      
      docTypeName = "ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (áõÕ.)"
      commentName = "²Ùë³ÃÇí- 29/05/17 N- " & docNum &" ¶áõÙ³ñ-               390.00 ²ñÅ.- 000 [Üáñ                 ]"
      
      ' Ստուգել որ Վճարման հանձնարարագիրն առկա է Հաճախորդի թղթապանակում
      status = CheckPayOrderAvailableOrNot(docTypeName, commentName)
      If Not status Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա չէ Հաճախորդի թղթապանակում")
            Exit Sub  
      End If
        
      ' Փակել հաճախորդի թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel_2").Close
      
      ' Վճարման հանձնարարագիրն ուղարկել հաստատման
      Call PaySys_Send_To_Verify()
       
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք VERIFIER  օգտագործողով
      Login("VERIFIER")
      Call ChangeWorkspace(c_Verifier1)
'      Call wTreeView.DblClickItem("|Ð³ëï³ïáÕ I ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
      BuiltIn.Delay(1000)
      Set verifyDocuments = New_VerificationDocument()
      verifyDocuments.User = "^A[Del]"
      Call GoToVerificationDocument("|Ð³ëï³ïáÕ I ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ", verifyDocuments)

      If Not wMDIClient.VBObject("frmPttel").Exists Then
            Log.Error("Հաստատվող փաստաթղթեր թղթապանակը չի բացվել")
            Exit Sub
      End If
      
      ' Պայամանագրի վավերացում
      colN = 3
      action = c_ToConfirm
      doActio = "Ð³ëï³ï»É"
      doNum = 1
      status =  ConfirmContractDoc(colN, docNum, action, doNum, doActio)
      
      If Not status Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա չէ")
            Exit Sub  
      End If
        
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
        
      ' Մուտք ARMSOFT  օգտագործողով
      Login("ARMSOFT")
      
      ' Մուտք Գլխավոր հաշվապահի ԱՇՏ   
      Call ChangeWorkspace(c_ChiefAcc)
      
      ' Մուտք Հաշվառված Վճարային փաստաթղթեր թղթապանակ
      workEnvName = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      workEnv = "Հաշվառված վճարային փաստաթղթեր"
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      
      If Not state Then
            Log.Error("Մուտք Հաշվառված վճարային փաստաթղթեր թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Պայմանագրի առկա լինելը ստուգող ֆունկցիա
      colN = 2
      status = CheckContractDoc(colN, docNum)
      
       If Not status Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա է Վճարային փաստաթղթեր թղթապանակում")
            Exit Sub  
      End If 
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
        
      ' Մուտք Արտքաին փոխանցումների ԱՇՏ  
      Call ChangeWorkspace(c_ExternalTransfers)
      
      ' Մուտք Ուղարկվող հանձնարարագրեր թղթապանակ
      workEnvName = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ"
      workEnv = "Ուղարկվող հանձնարարագրեր թղթապանակ"
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      
      If Not state Then
            Log.Error("Մուտք Ուղարկվող հանձնարարագրեր թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      Log.Message("SQL Check 2") 
     ' SQL ստուգում HI աղյուսակում
     queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '01' AND fOBJECT = '1630509' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' " & _ 
                              " AND fCURSUM = '390.00' AND fOP = 'TRF' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '390.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '01' AND fOBJECT = '425737041' AND fCUR = '000' AND fDATE = '2017-05-29 00:00:00' " & _ 
                              " AND fCURSUM = '390.00' AND fOP = 'TRF' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '390.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '01' AND fOBJECT = '1630442' AND fCUR = '000' AND fDATE = '2017-05-29 00:00:00'  " & _ 
                              " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '01' AND fOBJECT = '425737041' AND fCUR = '000' AND fDATE = '2017-05-29 00:00:00' " & _ 
                              " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '500.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
      
      ' Վճարման հանձնարարագիրն ուղարկել BankMail 
      colN = 2
      action = c_SendToBM
      doNum = 5
      doActio = "²Ûá"
      status = ConfirmContractDoc(colN, docNum, action, doNum, doActio)

      If Not status Then
            Log.Error("Վճարման հանձնարարագիրն առկա չէ ուղարկվող հանձնարարագրեր թղթապանակում")
            Exit Sub  
      End If       
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close

      ' Մուտք ուղարկված BankMail թղթապանակ
      workEnvName = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏí³Í  Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏí³Í BankMail"
      workEnv = "Ուղարկված BankMail "
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք Ուղարկված BankMail թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Պայմանագրի առկայության ստուգում 
      colN = 1
      status = CheckContractDoc(colN, docNum)
      If Not status Then
            Log.Error("Վճարման հանձնարարագիրն առկա չէ ուղարկված BankMail թղթապանակում")
            Exit Sub  
      End If 
      
      Log.Message("SQL Check 3") 
    ' SQL ստուգում HI աղյուսակում
     queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '01' AND fOBJECT = '1630509' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' " & _ 
                              " AND fCURSUM = '390.00' AND fOP = 'TRF' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '390.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '01' AND fOBJECT = '425737041' AND fCUR = '000' AND fDATE = '2017-05-29 00:00:00' " & _ 
                              " AND fCURSUM = '390.00' AND fOP = 'TRF' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '390.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '01' AND fOBJECT = '1630442' AND fCUR = '000' AND fDATE = '2017-05-29 00:00:00'  " & _ 
                              " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '01' AND fOBJECT = '425737041' AND fCUR = '000' AND fDATE = '2017-05-29 00:00:00' " & _ 
                              " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '500.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
              
      ' Մուտք BANKMAIL օգտագործողով
      Login("BANKMAIL")
       
      ' Մուտք ուղարկված փոխանցումներ  
      Call ChangeWorkspace(c_BM)

      ' Դիտել Վճարման հանձնարարագիրը 
      status = False
      Call WiewPayOrderFromTransferSent(wDate, fISN, childISN, status, wDateTime)
      
      If childISN =  "" Then
            Log.Error("Վճարման հանձնարարագիրն առկա չէ ուղարկվող թղթապանակում")
            Exit Sub  
      End If 
      
      ' Ծնող-զավակ կապի ստուգում  
      queryString = "SELECT fISN FROM DOCP WHERE fISN = " & childISN & " AND fPARENTISN = " & fISN
      wChildISN = Get_Query_Result(queryString)
      Log.Message(childISN)
      If  Trim(wChildISN) <> Trim(childISN) Then
            Log.Error("Ծնող զավակ կապի բացակայություն")
      End If
      
      colN = 2
      doNum = 5
      doActio = "²Ûá"
      status = Contract_To_Bank_Mail(colN, fISN)
      BuiltIn.Delay(3000)
      If Not status Then
            Log.Message("Պայմանագիրը չի գտնվել Bank Mail ուղարկվելու համար")
            Exit Sub
      End If
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Այսօրվա ամսաթվի ստացում
      tdDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y")
      
      ' Մուտք ուղարկված փոխանցումներ թղթապանակ
      workEnvName = "|BankMail ²Þî|öáË³ÝóáõÙÝ»ñ|àõÕ³ñÏí³Í ÷áË³ÝóáõÙÝ»ñ"
      workEnv = "Ուղարկված փոխանցումներ "
      state = AccessFolder(workEnvName, workEnv, stRekName, tdDate, endRekName, tdDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք Ուղարկված փոխանցումներ թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Պայմանագրի առկայության ստուգում
      colN = 3
      status = CheckContractDoc(colN, fISN)
      If Not status Then
          Log.Error("BankMail ուղարկված հաղորդագրությունը բացակայում է ուղարկված հաղորդագրություններ թղթապանակից")
      End If
      
      ' Պարամետրերի արժեքների ճշգրտում   
      paramName = "BMDBSERVER"
      paramValue = "qasql2017"
      Call  SetParameter(paramName, paramValue)
     
      sBody = ":20:" & fISN & vbCRLF _
                  & ":32A:" & wDate & "AMD390,"  & vbCRLF _
                  & ":50A:" & Replace(accDB, "/", "") & vbCRLF _ 
                  & payer &  vbCRLF _
                  & ":59:" & Replace(accCR, "/", "")& vbCRLF _ 
                  & receiver  & vbCRLF _
                  & ":70A:" & wAim         
      queryString = " SELECT Body FROM [qasql2017].BankMail_Testing.dbo.bmInterface WHERE AS_ISN = " & wChildISN
      bodyValue = Get_Query_Result(queryString)
      If  Trim(sBody) <> Trim(bodyValue) Then
        Log.Error("Փաստաթղթի տվյալները BankMail-ում չեն համապատասխանում dictionary-ով փոխանցվող տվյալների հետ")
      End If
         
      ' Մուտք ARMSOFT օգտագործողով
      Login("ARMSOFT")
          
      ' Մուտք ուղարկված BankMail թղթապանակ
      Call ChangeWorkspace(c_ExternalTransfers)

      ' Մուտք ուղարկված BankMail թղթապանակ
      workEnvName = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏí³Í  Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏí³Í BankMail"
      workEnv = "Ուղարկված BankMail "
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք Ուղարկված BankMail թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Պայմանագիրի համար մարել գործողության կատարում 
      colN = 1
      action = c_ToFade
      doNum  = 5
      doActio = "²Ûá"
      status = ConfirmContractDoc(colN, docNum, action, doNum, doActio)
      If Not status Then
            Log.Error("HT100 փաստաթուղթը առկա չէ ուղարկված BankMail թղթապանակում")
            Exit Sub
      ElseIf Not wMDIClient.WaitVBObject("frmASDocForm", 2000).Exists Then
            Log.Error("Խմբային հիշարար օրդեր պատուհաը չի բացվել")
            Exit Sub
      End If
      
      ' Խմբային հիշարար օրդերի  հաստատում 
      Call GroupReminderOrdersVer(grRemOrdISN, grRemOrdNum, wDate)
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք ուղարկված BankMail թղթապանակ
      workEnvName = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏí³Í  Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏí³Í BankMail"
      workEnv = "Ուղարկված BankMail "
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք Ուղարկված BankMail թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Պայմանագրի առկա չլինելը ստուգող ֆունկցիա  
      status = CheckContractDoc(colN, docNum)
      If status Then
          Log.Error("Փաստաթուղթը մարելուց հետո դեռ մնացել է ուղարկված BankMail թղթապանակում")
      End If
      
      Log.Message("SQL Check 4") 
    ' SQL ստուգում  
     queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fOBJECT = '1630509' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' " & _
                               " AND fCURSUM = '390.00' AND fOP = 'TRF' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '390.00' " & _
					                     " AND fBASEBRANCH = '00' AND fBASEDEPART = '1'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '425737041' AND fCUR = '000' AND fDATE = '2017-05-29 00:00:00' " & _
                                " AND fCURSUM = '390.00' AND fOP = 'TRF' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '390.00' " & _
					                      " AND fBASEBRANCH = '00' AND fBASEDEPART = '1' " 
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '1630442' AND fCUR = '000' AND fDATE = '2017-05-29 00:00:00' " & _
                                " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00' " & _
					                      " AND fBASEBRANCH = '00' AND fBASEDEPART = '1'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '425737041' AND fCUR = '000' AND fDATE = '2017-05-29 00:00:00' " & _
                                " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '500.00' " & _
					                      " AND fBASEBRANCH = '00' AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք Գլխավոր հաշվապահի ԱՇՏ
      Call ChangeWorkspace(c_ChiefAcc)
      
      ' Մուտք Հաշվառված վճարային փաստաթղթեր թղթապանակ
      workEnvName = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      workEnv = "Հաշվառված վճարային փաստաթղթեր "
      isnRekName = "DOCISN"
      wStatus = True
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, grRemOrdISN)
      If Not state Then
            Log.Error("Մուտք Հաշվառված վճարային փաստաթղթեր թղթապանակ ձախողվել է")
            Exit Sub
      End If
            
      ' Ստուգում որ Խմբային հիշարար օրդեր փաստաթուղթն առկա է Վճարային փաստաթղթեր թղթղպանակում
      If Not wMDIClient.WaitVBObject("frmPttel", 2000).Exists Then
            Log.Error("Վճարային փաստաթղթեր թղթապանակը չի բացվել")
            Exit Sub
      ElseIf wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount = 1 Then
            Log.Message("Խմբային հիշարար օրդեր փաստաթուղթն առկա է վճարային փաստաթղթեր թղթապանակում")
      End If
        
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք Հաշվառված վճարային փաստաթղթեր թղթապանակ
      workEnvName = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      workEnv = "Հաշվառված վճարային փաստաթղթեր "
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք Հաշվառված վճարային փաստաթղթեր թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Ստուգում որ Խմբային հիշարար օրդեր փաստաթուղթն առկա է 
      If Not wMDIClient.VBObject("frmPttel").Exists Then
            Log.Error("Վճարային փաստաթղթեր թղթապանակը չի բացվել")
            Exit Sub
      ElseIf wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
            Log.Error("Վճարման հանձնարարագրի փաստաթուղթն առկա չէ վճարային փաստաթղթեր թղթապանակում")
            Exit Sub
      End If
        
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Ծնող զավակ կապի ստուգում
      queryString = "SELECT fISN FROM DOCP WHERE fISN = " & grRemOrdISN & " AND fPARENTISN = " & fISN
      ordChildISN = Get_Query_Result(queryString)
      Log.Message(grRemOrdISN)
      If  Trim(ordChildISN) <> Trim(grRemOrdISN) Then
            Log.Error("Ծնող զավակ կապի բացակայություն")
      End If
      
      ' Փակել ծրագիրը
      Call Close_AsBank()  
End Sub
