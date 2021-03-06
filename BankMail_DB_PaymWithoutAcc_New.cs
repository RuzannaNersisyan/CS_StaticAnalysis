Option Explicit
'USEUNIT Akreditiv_Library
'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Online_PaySys_Library
'USEUNIT BankMail_Library
'USEUNIT BankMail_Library
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Library_Contracts

' Test Case ID 165077

Sub BankMail_DB_PaymWithoutAcc_New()
    
      Dim paramName, paramValue, confPath, confInput
      Dim ordType, fISN, wAcsBranch, wAcsDepart, wDate, docNum, cliCode, accDB, payer, ePayer, taxCods,_
              jurState, dbDropDown, coaNum, balAcc, accMask, accCur, accType, cliName, cCode, accNote, accNote2,_
              accNote3, acsBranch, acsDepart, acsType, pCardNum, socCard, accCR, receiver, eReceiver, summa, wCur,_
              wAim, jurStatR, bankCr, authorPerson, addInfo, wAddress, authPerson, rInfo              
      Dim workEnvName, workEnv, stRekName, endRekName, wStatus, isnRekName
      Dim queryString, sqlValue, colNum, sql_isEqual, state, value
      Dim wPayDate, paymentCode, status, rekvNum, docTypeName, commentName
      Dim colN, action, doNum, doActio, tdDate
      Dim childISN, wChildISN, wDateTime
      Dim sBody, bodyValue
      Dim startDate, fDate
        
      startDate = "20030101"
      fDate = "20250101"
      Call Initialize_AsBank("bank", startDate, fDate)
               
      ' Մուտք համակարգ ARMSOFT օգտագործողով
      Call Create_Connection()
      Login("ARMSOFT")
      
      ' Պարամետրերի արժեքների ճշգրտում   
      paramName = "BMUSEDB"
      paramValue = "1"
      Call  SetParameter(paramName, paramValue)
      
      paramName = "BMDBSERVER"
      paramValue = "qasql2017"
      Call  SetParameter(paramName, paramValue)
      
      paramName = "BMDBNAME"
      paramValue = "BankMail_Testing"
      Call  SetParameter(paramName, paramValue)
      
      ' Կարգավորումների ներմուծում
      confPath = "X:\Testing\Order confirm phases\NoVerify.txt"
      confInput = Input_Config(confPath)
      If Not confInput Then
          Log.Error("Կարգավորումները չեն ներմուծվել")
         Exit Sub
      End If
      
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ(Ընդլայնված)
      Call ChangeWorkspace(c_CustomerService)
      
      ' Մուտք աշխատանքային փաստաթղթեր
      workEnvName = "|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      workEnv = "Աշխատանքային փաստաթղթեր "
      wDate = "300717"
      stRekName = "PERN"
      endRekName = "PERK"
      wStatus = False
      state =  AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Սխալ՝ Աշխատանքային փաստաթղթեր թղթապանակ մուտք գործելիս")
            Exit Sub
      End If
      
      ' Վճարման հանձնարարագրի լրացում
      receiver = "Ð³Ï³é³Ï ´³Ûó³Ï"
      eReceiver = "But vice versa"
      summa = "357.1"
      wAim = "²¹³ßÇë"
      accDB = "77700/00001690100"
      bankCr = "11500"
      jurStatR = "21"
      ordType = "WOA"
      payer = "´³Ûó³Ï Ð³Ï³é³ÏÛ³Ý"
      dbDropDown = False
      Call PaymOrdToBeSentFill(ordType, fISN, wAcsBranch, wAcsDepart, wDate, docNum, cliCode, accDB, payer, ePayer, taxCods,_
                                                          jurState, dbDropDown, coaNum, balAcc, accMask, accCur, accType, cliName, cCode, accNote, accNote2,_
                                                          accNote3, acsBranch, acsDepart, acsType, pCardNum, socCard, accCR, receiver, eReceiver, summa, wCur,_
                                                          wAim, jurStatR, bankCr, authorPerson, addInfo, wAddress, authPerson, rInfo)
                                                          
      Log.Message("Փաստաթղթի ISN` " & fISN)
      Log.Message("Փաստաթղթի համարը՝ " & docNum)
      
      Log.Message("SQL Check 1") 
    ' SQL ստուգում HI աղյուսակում
     queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '11' AND fCUR = '000'   AND fCURSUM = '357.10' AND fOP = 'TRF' " & _ 
                              " AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '357.10'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '11' AND fCUR = '000'   AND fCURSUM = '357.10' AND fOP = 'TRF' " & _ 
                              " AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '357.10'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '11' AND fCUR = '000'   AND fCURSUM = '500.00' AND fOP = 'FEE' " & _
                               " AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '11' AND fCUR = '000'   AND fCURSUM = '500.00' AND fOP = 'FEE' " & _
                              " AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '500.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      ' Խմբագրել Վճարային հանձնարարագիրը
      Call ContractAction (c_ToEdit)
              
      If Not wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
            Log.Error("Վճարման հանձնարարագիրը չի բացվել")
            Exit Sub
      End If
      
      ' Անցում Լրացուցիչ բաժին
      Call GoTo_ChoosedTab(3)
      ' Վճարման օր դաշտի լրացում
      wPayDate = "010817"
      Call Rekvizit_Fill("Document", 3, "General", "PAYDATE", wPayDate)
      
      ' Անցում Վճարման տվյալներ բաժին
      Call GoTo_ChoosedTab(6)
      ' Հանձնարարականի կոդ դաշտի արժեքի ստուգում
      value = Get_Rekvizit_Value("Document", 6, "Mask", "PAYMENTCODE")
          
      If value <> "BKBK" Then
        Log.Error("Հանձնարարականի կոդ դաշտի արժեքը BKBK չէ:")
      Else
        Log.Message("Հանձնարարականի կոդ դաշտի արժեքի ստուգում - OK")
      End If
      
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(1, "Î³ï³ñ»É")
      
      Log.Message("SQL Check 2") 
    ' SQL ստուգում HI աղյուսակում
     queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '11' AND fCUR = '000'   AND fCURSUM = '357.10' AND fOP = 'TRF' " & _ 
                              " AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '357.10' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '11' AND fCUR = '000'   AND fCURSUM = '357.10' AND fOP = 'TRF' " & _ 
                              " AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '357.10' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '11' AND fCUR = '000'   AND fCURSUM = '500.00' AND fOP = 'FEE' " & _
                               " AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '11' AND fCUR = '000'   AND fCURSUM = '500.00' AND fOP = 'FEE' " & _
                              " AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '500.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      ' Ստուգում որ Վճարման հանձնարարագրի պայմանագիրն առկա է Հաճախորդներ թղթապանակում
      docTypeName = "ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (áõÕ.)"
      commentName = "²Ùë³ÃÇí- 30/07/17 N- "& docNum &" ¶áõÙ³ñ-               357.10 ²ñÅ.- 000 [ÊÙµ³·ñíáÕ           ]"
      state = CheckPayOrderAvailableOrNot(docTypeName, commentName)
      If Not state Then
            Log.Error("Վճարման հանձնարարագիրն առկա չէ Հաճախորդներ թղթապանակում")
            Exit Sub  
      End If
      
      ' Փաստաթղթի ուղարկում արտաքին բաժին
      Call PaySys_Send_To_External()
      
      ' Փակել հաճախորդի թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel_2").Close
      
      ' Փակել աշխատանքային փաստաթղթեր թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք Գլխավոր հաշվապահի ԱՇՏ   
      Call ChangeWorkspace(c_ChiefAcc)
      
      ' Մուտք Հաշվառված Վճարային փաստաթղթեր թղթապանակ
      workEnvName = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      workEnv = "Հաշվառված վճարային փաստաթղթեր"
      isnRekName = "DOCISN"
      wStatus = True
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Սխալ՝ Հաշվառված վճարային փաստաթղթեր թղթապանակ մուտք գործելիս")
            Exit Sub
      End If
      
      ' Փակել Վճարային փաստաթղթեր թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք Արտաքին փոխանցումների ԱՇՏ 
      Call ChangeWorkspace(c_ExternalTransfers)
      
      ' Մուտք Ուղարկվող հաձնարարագրեր թղթապանակ
      workEnvName = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ"
      workEnv = "Ուղարկվող հաձնարարագրեր"
      wStatus = False
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Սխալ՝ Ուղարկվող հաձնարարագրեր թղթապանակ մուտք գործելիս")
            Exit Sub
      End If
      
      Log.Message("SQL Check 3") 
    ' SQL ստուգում HI աղյուսակում
     queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fCUR = '000'   AND fCURSUM = '357.10' AND fOP = 'TRF' " & _ 
                               " AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '357.10' AND fBASEBRANCH = '00' And fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fCUR = '000'   AND fCURSUM = '357.10' AND fOP = 'TRF' " & _ 
                               " AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '357.10' AND fBASEBRANCH = '00' And fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fCUR = '000'   AND fCURSUM = '500.00' AND fOP = 'FEE'  " & _
                               " AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00' AND fBASEBRANCH = '00' And fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fCUR = '000'   AND fCURSUM = '500.00' AND fOP = 'FEE'  " & _
                               " AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '500.00' AND fBASEBRANCH = '00' And fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
      
      ' Վճարման հանձնարարագրի պայմանագիրն ուղարկել Bank Mail բաժին   
      colN = 2
      action = c_SendToBM
      doNum = 5
      doActio = "²Ûá"
      state = ConfirmContractDoc(colN, docNum, action, doNum, doActio)
      If Not state Then
            Log.Error("Վճարման հանձնարարագիրն առկա չէ Ուղարկվող հաձնարարագրեր թղթապանակում")
            Exit Sub  
      End If
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք Ուղարկված BankMail թղթապանակ
      workEnvName = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏí³Í  Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏí³Í BankMail"
      workEnv = "Ուղարկված BankMail"
      state = AccessFolder(workEnvName, workEnv, stRekName, wDate, endRekName, wDate, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Սխալ՝ Ուղարկված BankMail թղթապանակ մուտք գործելիս")
            Exit Sub  
      End If
      
      ' Ստուգում որ Վճարման հանձնարարագիրն առկա է
      colN = 1
      state = CheckContractDoc(colN, docNum)
      If Not state Then
            Log.Error("Վճարման հանձնարարագիրն առկա չէ Ուղարկված BankMail թղթապանակում")
            Exit Sub  
      End If
      
      Log.Message("SQL Check 4") 
    ' SQL ստուգում HI աղյուսակում
     queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fCUR = '000'   AND fCURSUM = '357.10' AND fOP = 'TRF' " & _ 
                               " AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '357.10' AND fBASEBRANCH = '00' And fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fCUR = '000'   AND fCURSUM = '357.10' AND fOP = 'TRF' " & _ 
                               " AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '357.10' AND fBASEBRANCH = '00' And fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fCUR = '000'   AND fCURSUM = '500.00' AND fOP = 'FEE'  " & _
                               " AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00' AND fBASEBRANCH = '00' And fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01' AND fCUR = '000'   AND fCURSUM = '500.00' AND fOP = 'FEE'  " & _
                               " AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '500.00' AND fBASEBRANCH = '00' And fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      ' Փակել ուղարկված BankMail թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք BANKMAIL օգտագործողով
      Login("BANKMAIL")
       
      ' Մուտք ուղարկված փոխանցումներ  
      Call ChangeWorkspace(c_BM) 

      ' Դիտել Վճարման հանձնարարագիրը
      status = True
      Call WiewPayOrderFromTransferSent(wDate, fISN, childISN, status, wDateTime)
      If  childISN = " " Then
            Log.Error("Վճարման հանձնարարագիրն առկա չէ")
            Exit Sub  
      End If
      
      ' Ծնող-զավակ կապի ստուգում  
      queryString = "SELECT fISN FROM DOCP WHERE  fPARENTISN = " & fISN
      wChildISN = Get_Query_Result(queryString)
      Log.Message("Զավակ փաստաթղթի ISN` " & childISN)
      If  Trim(wChildISN) <> Trim(childISN) Then
            Log.Error("Ծնող-զավակ կապի բացակայություն")
      End If
      
      ' Պայմանագրին ուղարկել BankMail
      colN = 2
      state = Contract_To_Bank_Mail(colN, fISN)
      If Not state Then
            Log.Message("Պայմանագիրը չի գտնվել Bank Mail ուղարկելու համար")
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
            Log.Error("Սխալ՝ Ուղարկված փոխանցումներ թղթապանակ մուտք գործելիս")
            Exit Sub  
      End If
      
      ' Պայմանագրի առկա լինելը ստուգող ֆունկցիա
      colN = 3
      state = CheckContractDoc(colN, fISN)
      If Not state Then
          Log.Error("BankMail ուղարկված հաղորդագրությունը բացակայում է ուղարկված հաղորդագրություններ թղթապանակից")
          Exit Sub
      End If
      
      sBody = ":20:" & fISN & vbCRLF _
                  & ":23E:BKBK" & vbCRLF _ 
                  & ":30A:" & tdDate & Replace(wDateTime,":","") & vbCRLF _
                  & ":32A:" & Replace(wDate,"/","") & "AMD357,1" & vbCRLF _
                  & ":23P:NP" & vbCRLF _
                  & ":50E:" & Replace(accDB, "/", "") & vbCRLF _ 
                  & "1/" & payer & vbCRLF _
                  & ":23Q:NP" & vbCRLF _          
                  & ":59F:" & bankCr & vbCRLF _
                  & "1/" & receiver & vbCRLF _
                  & ":70A:" & wAim   
          
      ' Տվյալների ստուգում BankMail_Testing.dbo.bmInterface  աղյուսակում
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