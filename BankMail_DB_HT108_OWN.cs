Option Explicit
'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Online_PaySys_Library
'USEUNIT BankMail_Library
'USEUNIT BankMail_Library
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Akreditiv_Library
'USEUNIT Library_Contracts

' Test Case ID 165071

Sub BankMail_DB_HT108_OWN()
      
      Dim ordType, fISN, wAcsBranch, wAcsDepart, payDate, docNum, cliCode, accDB, payer, ePayer, taxCods,_
              jurState, dbDropDown, coaNum, balAcc, accMask, accCur, accType, cliName, cCode, accNote, accNote2,_
              accNote3, acsBranch, acsDepart, acsType, pCardNum, socCard, accCR, receiver, eReceiver, summa, wCur,_
              wAim, jurStatR, bankCr, authorPerson, addInfo, wAddress, authPerson, rInfo 
      Dim colN, action, doNum, doActio, status, state, wMainForm
      Dim workEnvName, workEnv, stRekName, endRekName, wStatus, isnRekName
      Dim docTypeName, commentName, tcorrAcc, frmPttel
      Dim queryString, sqlValue, colNum, sql_isEqual
      Dim childISN, wChildISN
      Dim paramName, paramValue, sBody, bodyValue, wDateTime
      Dim wDate, tdDate, confPath, confInput
      Dim startDate, fDate, verifyDocuments

      startDate = "20030101"
      fDate = "20250101"
      Call Initialize_AsBank("bank", startDate, fDate)
               
      ' Մուտք համակարգ ARMSOFT օգտագործողով
      Call Create_Connection()
      Login("ARMSOFT")
      
      ' BMUSEDB պարամետրի արժեքը դնել 1
      paramName = "BMUSEDB"
      paramValue = "1"
      Call  SetParameter(paramName, paramValue)
      
      ' BMDBSERVER պարամետրի արժեքը դնել qasql2017
      paramName = "BMDBSERVER"
      paramValue = "qasql2017"
      Call  SetParameter(paramName, paramValue)
      
      ' BMDBNAME պարամետրի արժեքը դնել BankMail_Testing
      paramName = "BMDBNAME"
      paramValue = "BankMail_Testing"
      Call  SetParameter(paramName, paramValue)
      
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
            Log.Error("Մուտք Աշխատանքային փաստաթղթեր ձախողվել է")
            Exit Sub
      End If
      
      payDate = "290517"
      ordType = "BDG"
      accMask = "00057053311"
      wCur = "000"
      accType = "01"
      payer = "ê³ëáõÝóÇ ¸³íÇÃ"
      accCR = "90000/5001020"
      receiver = "Ð»ÕÝ³ñ ²ÕµÛáõñ"
      summa = "8000.00"
      wAim = "æñÇ í³ñÓ"
      dbDropDown = True
      accCur = "000"
      
      ' Վճարման հանձնարագրերի լրացում
      Call PaymOrdToBeSentFill(ordType, fISN, wAcsBranch, wAcsDepart, payDate, docNum, cliCode, accDB, payer, ePayer, taxCods,_
                                                          jurState, dbDropDown, coaNum, balAcc, accMask, accCur, accType, cliName, cCode, accNote, accNote2,_
                                                          accNote3, acsBranch, acsDepart, acsType, pCardNum, socCard, accCR, receiver, eReceiver, summa, wCur,_
                                                          wAim, jurStatR, bankCr, authorPerson, addInfo, wAddress, authPerson, rInfo)
      Log.Message(fISN)
      Log.Message(docNum) 
       
      Log.Message("SQL Check 1") 
      ' SQL ստուգում
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '11' AND fOBJECT = '1630509' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                                " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '8000.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '11' AND fOBJECT = '1714908' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                               " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '8000.00' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '11' AND fOBJECT = '1630353' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                               " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '500.00' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '11' AND fOBJECT = '1630442' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                                " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      ' Ստուգել որ Վճարման հանձնարարագրի պայմանագիրն առկա է հաճախորդի թղթապանակում
      docTypeName = "ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (áõÕ.)"
      commentName = "²Ùë³ÃÇí- 29/05/17 N- " & docNum &" ¶áõÙ³ñ-             8,000.00 ²ñÅ.- 000 [Üáñ                 ]"
      state = CheckPayOrderAvailableOrNot(docTypeName, commentName)
      If Not state Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա չէ հաճախորդի թղթապանակում")
            Exit Sub  
      End If
        
      ' Փակել հաճախորդի թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel_2").Close
      
      ' Փաստաթուղթն ուղարկել հաստատման
      colN = 2
      action = c_SendToVer
      doNum = 5
      doActio = "²Ûá"
      state = ConfirmContractDoc(colN, docNum, action, doNum, doActio)
      If Not state Then
            Log.Error("Պայմանագիրը չի ուղարկվել հաստատման")
            Exit Sub  
      End If
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք VERIFIER  օգտագործողով
      Login("VERIFIER")
      Call ChangeWorkspace(c_Verifier1)
'      Call wTreeView.DblClickItem("|Ð³ëï³ïáÕ I ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
      BuiltIn.Delay(1000)
      
      Set verifyDocuments = New_VerificationDocument()
      verifyDocuments.User = "^A[Del]"
      Call GoToVerificationDocument("|Ð³ëï³ïáÕ I ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ",verifyDocuments)
      If Not wMDIClient.VBObject("frmPttel").Exists Then
            Log.Error("Հաստատվող փաստաթղթեր թղթապանակը չի բացվել")
            Exit Sub
      End If
      
      ' Պայամանագրի վավերացում 
      colN = 3
      action = c_ToConfirm
      doActio = "Ð³ëï³ï»É"
      doNum = 1
      state = ConfirmContractDoc(colN, docNum, action, doNum, doActio)
      If Not state Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա չէ Հաստատվող վճարային փաստաթղթեր թղթապանակում և չի վավերացվել")
            Exit Sub  
      End If
        
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք ARMSOFT  օգտագործողով
      Login("ARMSOFT")
      
      ' Մուտք Գլխավոր հաշվապահի ԱՇՏ   
      Call ChangeWorkspace(c_ChiefAcc)
      
      ' Մուտք Հաշվառված վճարային փաստաթղթեր թղթապանակ
      workEnvName = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      workEnv = "Հաշվառված վճարային փաստաթղթեր "
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք Հաշվառված վճարային փաստաթղթեր թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Պայմանագրի առկա լինելը ստուգող ֆունկցիա
      colN = 2
      state = CheckContractDoc(colN, docNum)
      If Not state Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա է Վճարային փաստաթղթեր թղթապանակում")
            Exit Sub  
      End If 
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
        
      ' Մուտք Արտաքին փոխանցումների ԱՇՏ/ Ուղարկվող հանձնարարագրեր  
      Call ChangeWorkspace(c_ExternalTransfers)
      
      ' Մուտք Ուղարկվող հանձնարարագրեր թղթապանակ
      workEnvName = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ"
      workEnv = "Ուղարկվող հանձնարարագրեր "
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք Ուղարկվող հանձնարարագրեր թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      Log.Message("SQL Check 2") 
      ' SQL ստուգում
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '1630509' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                                " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '8000.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fOBJECT = '1714908' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                               " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '8000.00'  "& _ 
                               " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fOBJECT = '1630353' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                               " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '500.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '1630442' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                                " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
      
      ' Վճարման հանձնարարագրի պայմանագիրն ուղարկել BankMail 
      colN = 2
      action = c_SendToBM
      doNum = 5
      doActio = "²Ûá"
      state = ConfirmContractDoc(colN, docNum, action, doNum, doActio)
      If Not state Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա չէ")
            Exit Sub  
      End If 
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք ուղարկված BankMail թղթապանակ
      workEnvName = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏí³Í  Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏí³Í BankMail"
      workEnv = "ուղարկված BankMail "
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք ուղարկված BankMail թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Պայմանագրի առկա լինելը ստուգող ֆունկցիա 
      colN = 1
      state = CheckContractDoc(colN, docNum)
      If Not state Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա չէ ուղարկված BankMail թղթապանակում")
            Exit Sub  
      End If
      
      Log.Message("SQL Check 3") 
      ' SQL ստուգում 
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '1630509' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                                " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '8000.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fOBJECT = '1714908' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                               " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '8000.00'"& _ 
                               " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fOBJECT = '1630353' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                               " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'D' AND fSUID = '77' AND fSUM  = '500.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '1630442' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                                " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
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
         
      ' Դիտել Վճարման հանձնարարագրի պայմանագիրը 
      status = False
      Call WiewPayOrderFromTransferSent(wDate, fISN, childISN, status, wDateTime)
      If childISN =  "" Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա չէ Ուղարկվող թղթապանակում")
            Exit Sub  
      End If 
      
      ' Ծնող-զավակ կապի ստուգում  
      queryString = "SELECT fISN FROM DOCP WHERE fISN = " & childISN & " AND fPARENTISN = " & fISN
      wChildISN = Get_Query_Result(queryString)
      Log.Message(childISN)
      If  Trim(wChildISN) <> Trim(childISN) Then
            Log.Error("Ծնող-զավակ կապի բացակայություն")
      End If
      
      ' Պայմանագիրն ուղարկել մասնակի խմբագրման
      colN = 2
      action = c_SendToPartEd
      doNum = 2
      doActio = "Î³ï³ñ»É"
      state = ConfirmContractDoc(colN, fISN, action, doNum, doActio)
      BuiltIn.Delay(3000)
      If Not state Then
            Log.Error("Պայմանագիրը չի գտնվել մասնակի խմբագրման ուղարկելու համար")
            Exit Sub
      End If

      ' Մուտք համակարգ ARMSOFT օգտագործողով  
      Call Create_Connection()
      Login("ARMSOFT")
      
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ(Ընդլայնված) ԱՇՏ
      Call ChangeWorkspace(c_CustomerService)
      
      ' Մուտք աշխատանքային փաստաթղթեր թղթապանակ
      workEnvName = "|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      workEnv = "աշխատանքային փաստաթղթեր "
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք աշխատանքային փաստաթղթեր թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      Log.Message("SQL Check 4") 
      ' SQL ստուգում   
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '1630509' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                                " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '8000.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fOBJECT = '1714908' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                               " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'D' AND fSUID = '77' AND fSUM =  '8000.00' "& _ 
                               " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*)  FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fOBJECT = '1630353' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                               " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'D' AND fSUID = '77' AND fSUM =  '500.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '1630442' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                                " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'C' AND fSUID = '77' AND fSUM =  '500.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
      
      ' Փաստաթղթի առկայության ստուգում Աշխատանքային փաստաթղթեր թղթապանակում
      colN = 2
      state = CheckContractDoc(colN, docNum)
      BuiltIn.Delay(3000)
      If Not state Then
            Log.Error("Պայմանագիրը բացակայում է Աշխատանքային փաստաթղթեր թղթապանակից")
            Exit Sub
      End If
      
      ' Խմբագրել Վճարման հանձնարարագիրը
      action = c_ToEdit
      Call ContractAction(action)
      BuiltIn.Delay(2000)
      
      ' Հաշիվ դեբետ դաշտի արժեքի ստացում
      accDB = wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("ASAmACC").VBObject("TDBBank").Text
      ' Տարանցիկ հաշիվ դաշտի լրացում   
      tcorrAcc = "000411200"
      Call Rekvizit_Fill("Document", 2, "General", "TCORRACC", tcorrAcc)
      ' Հաստատել կոճակի սեղմում
      Call ClickCmdButton(1, "Î³ï³ñ»É")
      
      Log.Message("SQL Check 5") 
      ' SQL ստուգում   
      queryString = " SELECT COUNT(*)  FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '1630279' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                                " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '8000.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*)  FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fOBJECT = '1714908' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                               " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '8000.00' "& _ 
                               " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*)  FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fOBJECT = '1630353' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                               " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'D' AND fSUID = '77' AND fSUM =  '500.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '1630442' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                                " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1 
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      ' Փակել աշխատանքային փաստաթղթեր թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք Արտաքին փոխանցումների ԱՇՏ  
      Call ChangeWorkspace(c_ExternalTransfers)
      
      ' Մուտք Մասնակի խմբագրվող հանձնարարագրեր թղթապանակ
      workEnvName = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|Ø³ëÝ³ÏÇ ËÙµ³·ñíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ"
      workEnv = "Մասնակի խմբագրվող հանձնարարագրեր "
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      
      If Not state Then
            Log.Error("Մուտք Մասնակի խմբագրվող հանձնարարագրեր թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Պայմանագրն ուղարկել հաստատման
      colN = 1
      action = c_SendToVer
      doNum = 5
      doActio = "²Ûá"
      state = ConfirmContractDoc(colN, docNum, action, doNum, doActio)
      
      If Not state Then
           Log.Error("Պայմանագիրը չի գտնվել մասնակի խմբագրվող հանձնարարագրեր թղթապանակում և չի ուղարկվել հաստատման ")
           Exit Sub
      End If
      
      ' փակել մասնակի խմբագրվող հանձնարարագրեր թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      Login("VERIFIER")
      Call wTreeView.DblClickItem("|Ð³ëï³ïáÕ I ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
     
      Set verifyDocuments = New_VerificationDocument()
      verifyDocuments.User = "^A[Del]"
      Call GoToVerificationDocument("|Ð³ëï³ïáÕ I ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ",verifyDocuments)

      ' Վավերացնել պայմանագիրը
      colN = 3
      action = c_ToConfirm
      doNum = 1
      doActio = "Ð³ëï³ï»É"
      state =  ConfirmContractDoc(colN, docNum, action, doNum, doActio)
      
      If Not state Then
           Log.Error("Պայմանագիրը չի գտնվել մասնակի խմբագրվող հանձնարարագրեր թղթապանակում և չի վավերացվել ")
           Exit Sub
      End If
      
      ' Փակել հաստատվող փաստաթղթեր թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
       
      ' Մուտք Գլխավոր հաշվապահի ԱՇՏ 
      Login("ARMSOFT") 
      Call ChangeWorkspace(c_ChiefAcc)
      
      ' Մուտք Վճարային փաստաթղթեր թղթապանակ
      workEnvName = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      workEnv = "Վճարային փաստաթղթեր"
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք Վճարային փաստաթղթեր թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Պայմանագրի առկայության ստուգում
      colN = 2
      state = CheckContractDoc(colN, docNum)
      If Not state Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա չէ Հաշվառված վճարային փաստաթղթեր թղթապանակում")
            Exit Sub  
      End If 
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
        
      ' Մուտք Արտաքին փոխանցումների ԱՇՏ  
      Call ChangeWorkspace(c_ExternalTransfers)
      
      ' Մուտք Ուղարկված BankMail թղթապանակ
      workEnvName = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ"
      workEnv = "Ուղարկվող հաձնարարգրեր"
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք Ուղարկվող հաձնարարագրեր թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      Log.Message("SQL Check 6") 
     ' SQL ստուգում   
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '1630279' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                                " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '8000.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fOBJECT = '1714908' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                               " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '8000.00' "& _ 
                               " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fOBJECT = '1630353' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                               " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '500.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '1630442' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                                " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
      
      ' Վճարման հանձնարարագրի պայմանագիրն ուղարկել BankMail 
      colN = 2
      action = c_SendToBM
      doNum = 5
      doActio = "²Ûá"
      state = ConfirmContractDoc(colN, docNum, action, doNum, doActio)
      If Not state Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա չէ ուղարկվող հանձնարարագրեր թղթապանակում")
            Exit Sub  
      End If 
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք Ուղարկված BankMail թղթապանակ
      workEnvName = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏí³Í  Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏí³Í BankMail"
      workEnv = "Ուղարկված BankMail"
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք Ուղարկված BankMail թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Պայմանագրի առկա լինելը ստուգող ֆունկցիա
      colN = 1
      state = CheckContractDoc(colN, docNum)
      If Not state Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա չէ ուղարկված BankMail թղթապանակում")
            Exit Sub  
      End If 
      
      Log.Message("SQL Check 7") 
    ' SQL ստուգում  
      queryString = " SELECT COUNT(*)  FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '1630279' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                                " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '8000.00'  "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*)  FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fOBJECT = '1714908' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _ 
                               " AND fCURSUM = '8000.00' AND fOP = 'TRF' AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '8000.00'"& _ 
                               " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fOBJECT = '1630353' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                               " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'D' AND fSUID = '77' AND fSUM =  '500.00' "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                                " AND fTYPE = '01' AND fOBJECT = '1630442' AND fCUR = '000'  AND fDATE = '2017-05-29 00:00:00' "& _
                                " AND fCURSUM = '500.00' AND fOP = 'FEE' AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00'  "& _ 
                                " AND fBASEBRANCH = '00'  AND fBASEDEPART = '1' "
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

      ' Դիտել Վճարման հանձնարարագրի պայմանագիրն
      status = True
      Call WiewPayOrderFromTransferSent(wDate, fISN, childISN, status, wDateTime)
      If childISN =  "" Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա չէ Ուղարկվող թղթապանակում")
            Exit Sub  
      End If 
      
      ' Ծնող-զավակ կապի ստուգում  
      queryString = "SELECT fISN FROM DOCP WHERE fISN = " & childISN & " AND fPARENTISN = " & fISN
      wChildISN = Get_Query_Result(queryString)
      Log.Message(childISN)
      If  Trim(wChildISN) <> Trim(childISN) Then
            Log.Error("Ծնող-զավակ կապի բացակայություն")
      End If
      
      ' Պայմանագրին ուղարկել BankMail
      colN = 2
      state = Contract_To_Bank_Mail(colN, fISN)
      BuiltIn.Delay(3000)
      If Not state Then
            Log.Error("Պայմանագիրը չի գտնվել BankMail ուղարկելու համար")
            Exit Sub
      End If
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Այսօրվա ամսաթվի ստացում
      tdDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y")
      
      ' Մուտք ուղարկված փոխանցումներ թղթապանակ
      workEnvName = "|BankMail ²Þî|öáË³ÝóáõÙÝ»ñ|àõÕ³ñÏí³Í ÷áË³ÝóáõÙÝ»ñ"
      workEnv = "Ուղարկված փոխանցումներ"
      state = AccessFolder(workEnvName, workEnv, stRekName, tdDate, endRekName, tdDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Մուտք Ուղարկված փոխանցումներ թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Պայմանագրի առկայության ստուգում
      colN = 3
      state = CheckContractDoc(colN, fISN)
      If Not state Then
          Log.Error("BankMail ուղարկված հաղորդագրությունը բացակայում է ուղարկված հաղորդագրություններ թղթապանակից")
          ExiT Sub
      End If
      
      sBody = ":20:" & fISN & vbCRLF _
                  & ":23E:TUBG" & vbCRLF _ 
                  & ":30A:" & tdDate & Replace(wDateTime,":","") & vbCRLF _
                  & ":32A:" & wDate & "AMD8000," & vbCRLF _
                  & ":23P:NP" & vbCRLF _
                  & ":50E:" & Replace(accDB, "/", "") & vbCRLF _ 
                  & "1/" & payer & vbCRLF _
                  & "3/AB123456" & vbCRLF _
                  & "8/ºñ»í³Ý ºñ»í³ÝÛ³Ý ÷" & vbCRLF _ 
                  & ":23Q:LP" & vbCRLF _          
                  & ":59E:" & Replace(accCR, "/", "") & vbCRLF _
                  & "1/" & receiver & vbCRLF _
                  & ":70A:" & wAim & vbCRLF _
                  & ":77B:PTD/OTM000000E000000OT/0/0"
      
      ' Տվյալների ստուգում [qasql2017].BankMail_Testing.dbo.bmInterface աղյուսակում
      queryString = " SELECT Body FROM [qasql2017].BankMail_Testing.dbo.bmInterface WHERE AS_ISN = " & wChildISN
      bodyValue = Get_Query_Result(queryString)
      If  Trim(sBody) <> Trim(bodyValue) Then
            Log.Error("Փաստաթղթի տվյալները BankMail-ում չեն համապատասխանում dictionary-ով փոխանցվող տվյալների հետ")
      End If
         
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Փակել ծրագիրը
      Call Close_AsBank()
End Sub