Option Explicit
'USEUNIT  Library_Common
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT Debit_Dept_Library
'USEUNIT Clients_Library
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Subsystems_Special_Library
'USEUNIT Payment_Except_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Mortgage_Library
' Test ID 132302
 
Sub Bill_Receivables_BalanceSheet_Test()

      Dim direction, wColNum, contType, fISN, cliCode, accAcc, comment, wDate, acsBranch, branchSect, acsType,_
               autoDebt, useAccBalanc, accConnect, headNum, autoDateChild, typeAutoDate, fixedDays, agrPeriod,_
               agrPeryodDay, passDirrect, passType, dateAgr, clsDays, state, brType, notClass, subjRisk,_
               sector, wAim, riksDegree, repCode, wNote, wNote2, wNote3, pprCode, dateClose, cenceled, n16AccType, _
               fillAccs, complRef, status, storageAcc, cost, income, accOutAgr 
              
      Dim folderDirect, folderName, rekvName, debtLetISN, storeISN, fDATE, sDATE
      
      Dim confPath, confInput, risk, perc, queryString, sqlValue, ColNum, sql_isEqual
      
      Dim Prov, Removal, sumRes, sumUnres, wComment, acsSect, sumAgr, action, wSumma
      
      Dim rChng, rDeal, fileName1, savePath, fName, fileName2, param, docName
       
      Dim fillAcsBranch, fillAcsDepart, fillAcsType, acsDepart, fillDefault, tdbgView
      
      Dim wDbt, wRes, wOut, wInc, wCls, wRsk, paramN, wFrmPttel, colN, riskClassfISN 
      
      Dim  docTypeName, dbtISN, sDatePar, dateGive, eDatePar, datePar, dateCl, wBranch, wDepart, wAcsType
            
      fDATE = "20250101"
      sDATE = "20120101"
      Call Initialize_AsBank("bank", sDATE, fDATE)
      
      ' Մուտք գործել համակարգ ARMSOFT օգտագործողով 
      Call Create_Connection()
      Login("ARMSOFT")
      
      ' Մուտք Դեբիտորական պարտքեր ԱՇՏ
      Call ChangeWorkspace(c_BillReceivables)

      ' Դեբիտորական պարտք փաստաթղթի ստեղծում
      direction = "|¸»µÇïáñ³Ï³Ý å³ñïù»ñ|Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ"
      contType = "¸»µÇïáñ³Ï³Ý å³ñïù(ºïÑ³ßí»Ïßé³ÛÇÝ)"
      status = True
      wColNum = 1 
      state = False
      cliCode = "8101800/000007"  
      wDate = "010120"
      acsBranch = "00"
      branchSect = "1"
      acsType = "BR0"
      autoDateChild = 1
      typeAutoDate = "1"
      fixedDays = "15"
      passDirrect = "0"
      brType = "9" 
      sector = "E"
      wAim = "00"
      riksDegree = "0.00"
      repCode = "103"
      fillAccs = 1
      complRef = 1
      storageAcc = "19510066600"
      cost = "001344400"
      income = "000441900"  
      Call SelectContracType(direction, wColNum, contType, fISN, cliCode, cliCode, comment, wDate, acsBranch, branchSect, acsType,_
                                            autoDebt, useAccBalanc, accConnect, headNum, autoDateChild, typeAutoDate, fixedDays, agrPeriod,_
                                            agrPeryodDay, passDirrect, passType, dateAgr, clsDays, state, brType, notClass, subjRisk,_
                                            sector, wAim, riksDegree, repCode, wNote, wNote2, wNote3, pprCode, dateClose, cenceled, n16AccType, _
                                            fillAccs, complRef, status, storageAcc, cost, income, accOutAgr)
      Log.Message("Դեբիտորական պարտք փաստաթղթի fISN` " & fISN)
                                            
       ' Մուտք աշխատանքային փաստաթղթեր թղթապանակ և  պայմանագրի առկայության ստուգում
      folderDirect = "|¸»µÇïáñ³Ï³Ý å³ñïù»ñ|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      folderName = "Աշխատանքային փաստաթղթեր"
      rekvName = "NUM"
      state =  OpenFolder(folderDirect, folderName, rekvName, cliCode)
                           
      If Not state Then
            Log.Error("Դեբիտորական պարտք փաստաթուղթն առկա չէ Աշխատանքային փաստաթղթեր թղթապանակում")
            Exit Sub
      End If
      
                  'CONTRACTS
                  queryString = " SELECT COUNT(*) FROM CONTRACTS WHERE fDGISN = " & fISN & _
                                            " and fDGSUMMA = '0.00' and fDGALLSUMMA = '0.00' " & _
                                            " and fDGRISKDEGNB = '0.00' and fDGRISKDEGREE = '0.00' " & _
                                            " and fDGCUR = '001' and fDGMPERCENTAGE = '0.00' "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'DOCS
                  queryString = " Select COUNT(*) from DOCS where fISN = " & fISN & " and fNAME = 'BRBillNB' and fBODY = '" & vbCRLF _
                                       & "CODE:8101800/000007" & vbCRLF _
                                       & "CLICOD:00000093" & vbCRLF _
                                       & "ACCTYPE:2" & vbCRLF _
                                       & "NAME:§Î³ÛÍ³Ï¦ êäÀ" & vbCRLF _
                                       & "CURRENCY:001" & vbCRLF _
                                       & "DATE:20200101" & vbCRLF _
                                       & "ACSBRANCH:00" & vbCRLF _
                                       & "ACSDEPART:1" & vbCRLF _
                                       & "ACSTYPE:BR0" & vbCRLF _
                                       & "JURSTAT:11" & vbCRLF _
                                       & "VOLORT:9X" & vbCRLF _
                                       & "PETBUJ:2" & vbCRLF _
                                       & "REZ:1" & vbCRLF _
                                       & "RELBANK:0" & vbCRLF _
                                       & "RABBANK:0" & vbCRLF _
                                       & "AUTODATECHILD:1" & vbCRLF _
                                       & "TYPEAUTODATE:1" & vbCRLF _
                                       & "FIXEDDAYS:15" & vbCRLF _
                                       & "PASSOVDIRECTION:0" & vbCRLF _
                                       & "SECTOR:E" & vbCRLF _
                                       & "AIM:00" & vbCRLF _
                                       & "PERRES:1" & vbCRLF _
                                       & "REPCODE:103" & vbCRLF _
                                       & "SUBJRISK:0" & vbCRLF _
                                       & "FILLACCS:0" & vbCRLF _
                                       & "OPENACCS:0" & vbCRLF _
                                       & "'" 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                        
      ' Փաստաթուղթն ուղարկել հաստատման
      Call PaySys_Send_To_Verify()
      Call Close_Pttel("frmPttel")
      
                  'HIF
                  queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & fISN & _
                                            " and ((fSUM = '0.00' and fCURSUM = '0.00' and fOP = 'ORC' and fTRANS = '1') " & _
                                            " or (fSUM = '1.00' and fCURSUM = '0.00' and fOP = 'PRS' and fTRANS = '1') " & _
                                            " or (fSUM = '0.00' and fCURSUM = '0.00' and fOP = 'RSK' and fTRANS = '1')) "
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
      ' Մուտք Պայմանագրեր թղթապանակ և  պայմանագրի առկայության ստուգում
      folderDirect = "|¸»µÇïáñ³Ï³Ý å³ñïù»ñ|ä³ÛÙ³Ý³·ñ»ñ"
      folderName = "Պայմանագրեր"
      rekvName = "ACCBAL"
      cliCode = "8101800"
      state =  OpenFolder(folderDirect, folderName, rekvName, cliCode)
                           
      If Not state Then
            Log.Error("Դեբիտորական պարտք փաստաթուղթն առկա չէ Պայմանագրեր թղթապանակում")
            Exit Sub
      End If
      
      ' Ռիսկի դասիչ և պահուստավորում գործողության կատարում
      risk = "05"
      perc = "100"
      Call FillDoc_Risk_Classifier(wDate, risk, perc)
      
      wFrmPttel = "frmPttel_2"
      paramN = c_ViewEdit & "|" & c_Risking & "|" & c_RisksPersRes
      colN = 0
      cliCode = "8101800/000007"  
      Call GetfISNFromActionsView(paramN, wDate, wDate, wFrmPttel, colN, cliCode, riskClassfISN )
      Log.Message(" Ռիսկի դասիչ և պահուստավորման տոկոս փաստաթղթի ISN` " & riskClassfISN)

                 'HI
                  queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & riskClassfISN & _
                                            " and fTYPE = 'N0' and fCURSUM = '0.00' " & _
                                            " and ((fSUM = '100.00' and fOP = 'PRS')" & _
                                            " or (fSUM = '0.00' and fOP = 'RSK')) " 
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'DOCS
                  queryString = " Select COUNT(*) from DOCS where fISN = " & riskClassfISN & " and fNAME = 'BRTSRsPr' and fBODY = '" & vbCRLF _
                                       & "CODE:8101800/000007" & vbCRLF _
                                       & "DATE:20200101" & vbCRLF _
                                       & "RISK:05" & vbCRLF _
                                       & "PERRES:100" & vbCRLF _
                                       & "COMMENT:èÇëÏÇ ¹³ëÇã ¨ å³Ñáõëï³íáñÙ³Ý ïáÏáë" & vbCRLF _
                                       & "USERID:  77" & vbCRLF _
                                       & "'" 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
      
      ' Պահուստավորում գործողության կատարում
      wDate = "150120"
      action = c_Store
      state = False
      sumRes = "100000"
      sumUnres = "0.00"
      wComment = "ä³Ñáõëï³íáñáõÙ"
      Call ProvisionAction(action, storeISN, wDate, state, wSumma, sumAgr, sumRes, sumUnres, wComment, acsBranch, acsSect)
      Log.Message("Պահուստավորում փաստաթղթի ISN` " & storeISN)
      
                  'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & storeISN & _
                                            " and fSUM = '100000.00' and fCURSUM = '100000.00' and fOP = 'RST' and fCUR = '000' " & _
                                            " and fTYPE = '01' and (fDBCR = 'C' or fDBCR = 'D') " 
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  'HIR
                  queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & storeISN &_
                                           " and fCURSUM = '100000.00' and fOP = 'RES' and fCUR = '000' "&_
                                           " and fTYPE = 'R4' and  fDBCR = 'D'" 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                            " and fTYPE = 'R4' and fLASTREM = '100000.00' "& _
                                            " and fPENULTREM = '0.00' and fSTARTREM = '0.00' "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

      ' Ընդհանուր դիտում 
      rChng = 1
      rDeal = 1
	    datePar = "RDATE"
      Call OverallView(datePar, wDate, rChng, rDeal)

      ' Ֆայլի ջնջում
      fileName1 = "\\host2\Sys\Testing\DebitDebt\ActualBillReceiv.txt"
      aqFile.Delete(fileName1)
      
      savePath = "\\host2\Sys\Testing\DebitDebt\"
      fName = "ActualBillReceiv.txt"
      fileName2 = "\\host2\Sys\Testing\DebitDebt\ExpectlBillReceiv.txt"

      ' Հիշել քաղվածքը
      Call SaveDoc(savePath, fName)

      ' Ամսաթվերի բացառում
      param = "([0-9]{2}[/][0-9]{2}[/][0-9]{2}).[0-9] [0-9]{2}[:][0-9]{2}|([0-9]{2}[/][0-9]{2}[/][0-9]{2})|(<td class=""statement-trxn-docnum table-cell"">[0][0][0]([0-9]{3})<[/]td>)"
      
      ' Ֆայլերի համեմատում 
      Call Compare_Files(fileName2, fileName1, param)
      
      Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("FrmSpr").Close
      
      ' Փոխել Գրասենյակը|Բաժինը|Տիպը գործողության կատարում
      acsBranch = "01"
      acsDepart = "2"
      fillAcsBranch = 1
      fillAcsDepart = 1
      fillAcsType = 1
      status = True
      Call ChangeBranchDepartType(status, fillAcsBranch, fillAcsDepart, fillAcsType, acsBranch, acsDepart, acsType, fillDefault )
      
      ' Ստուգել Գրասենյակ և Բաժին դաշտերի արժեքները
       Set tdbgView = wMDIClient.VBObject("frmPttel").VBObject("TDBGView")
      
      wBranch = wMDIClient.VBObject("frmPttel").GetColumnIndex("ACSBRANCH")
      wDepart = wMDIClient.VBObject("frmPttel").GetColumnIndex("ACSDEPART")
      wAcsType = wMDIClient.VBObject("frmPttel").GetColumnIndex("ACSTYPE")
      If  Not (Trim( tdbgView.Columns.Item(wBranch).Value) = Trim(acsBranch)  and   Trim( tdbgView.Columns.Item(wDepart).Value) = Trim(acsDepart)  and Trim( tdbgView.Columns.Item(wAcsType).Value) = Trim(acsType))  Then
            Log.Error("Փոխել Գրասենյակը|Բաժինը|Տիպը գործողության կատարումը հաջողությամբ չի իրականացել")
      End If
      
                   'HIF
                  queryString = " select COUNT(*) from HIF where fBASE  = " & fISN &_
                                           "  and fTYPE = 'N0' and fCURSUM = '0.00' and fADB = '0'" &_
                                           " and (fSUM = '0.00' or fSUM = '1.00') " &_
                                           " and (fOP = 'ORC' or fOP = 'PRS' or fOP = 'RSK')" 
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " Select COUNT(*) from DOCS where fISN = " & fISN & " and fNAME = 'BRBillNB' and fBODY = '"& vbCRLF _
                                           & "CODE:8101800/000007"& vbCRLF _
                                           & "CLICOD:00000093"& vbCRLF _
                                           & "ACCTYPE:2"& vbCRLF _
                                           & "NAME:§Î³ÛÍ³Ï¦ êäÀ"& vbCRLF _
                                           & "CURRENCY:001"& vbCRLF _
                                           & "DATE:20200101"& vbCRLF _
                                           & "ACSBRANCH:01"& vbCRLF _
                                           & "ACSDEPART:2"& vbCRLF _
                                           & "ACSTYPE:BR0"& vbCRLF _
                                           & "JURSTAT:11"& vbCRLF _
                                           & "VOLORT:9X"& vbCRLF _
                                           & "PETBUJ:2"& vbCRLF _
                                           & "REZ:1"& vbCRLF _
                                           & "RELBANK:0"& vbCRLF _
                                           & "RABBANK:0"& vbCRLF _
                                           & "AUTODATECHILD:1"& vbCRLF _
                                           & "TYPEAUTODATE:1"& vbCRLF _
                                           & "FIXEDDAYS:15"& vbCRLF _
                                           & "PASSOVDIRECTION:0"& vbCRLF _
                                           & "SECTOR:E"& vbCRLF _
                                           & "AIM:00"& vbCRLF _
                                           & "PERRES:1"& vbCRLF _
                                           & "REPCODE:103"& vbCRLF _
                                           & "SUBJRISK:0"& vbCRLF _
                                           & "FILLACCS:0"& vbCRLF _
                                           & "OPENACCS:0"& vbCRLF _
                                           & "'"
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
      
      ' Ապապահուստավորում գործողության կատարում
      action = c_Store
      wDate = "010220"
      sumRes = "0.00"
      sumUnres = "50000.00"
      wComment = "²åå³å³Ñáõëï³íáñáõÙ"
      Call ProvisionAction(action, storeISN, wDate, state, wSumma, sumAgr, sumRes, sumUnres, wComment, acsBranch, acsSect)
      Log.Message("Ապապահուստավորում փաստաթղթի ISN` " & storeISN)

                  'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & storeISN & _
                                            " and fSUM = '50000.00' and fCURSUM = '50000.00' and fOP = 'RST' and fCUR = '000'" &_
                                            " and fTYPE = '01' and (fDBCR = 'C' or fDBCR = 'D')"
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIR
                  queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & storeISN & _
                                            " and fCURSUM = '50000.00' and fOP = 'UNR' and fCUR = '000'" &_
                                            " and fTYPE = 'R4' and  fDBCR = 'C'"
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                            " and fTYPE = 'R4' and fLASTREM = '50000.00' "& _
                                            " and fPENULTREM = '100000.00' and fSTARTREM = '0.00' "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

      ' Խմբային պահուստավորում գործողության կատարում
      wRes = 1
      wDate = "150220"
      Call GroupCalculation(wDate, wDate, wDbt, wRes, wOut, wInc, wCls, wRsk)

      paramN = c_OpersView
      colN = 5
      docTypeName = "ä³Ñáõëï³íáñáõÙ"
      Call GetfISNFromActionsView(paramN, wDate, wDate, wFrmPttel, colN, docTypeName, dbtISN )
      log.Message("Պահուստավորում փաստաթղթի ISN` " & dbtISN)
      
                  'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & dbtISN & _
                                            " and fSUM = '1950000.00' and fCURSUM = '1950000.00' and fOP = 'RST' and fCUR = '000'" &_
                                            " and fTYPE = '01' and (fDBCR = 'C' or fDBCR = 'D') "
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIR
                  queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & dbtISN & _
                                            " and fCURSUM = '1950000.00' and fOP = 'RES' and fCUR = '000'" &_
                                            " and fTYPE = 'R4' and  fDBCR = 'D'"
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                            " and fTYPE = 'R4' and fLASTREM = '2000000.00'"& _
                                            " and fPENULTREM = '50000.00' and fSTARTREM = '0.00' "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

      ' Պայմանագրի փակում
      wDate = "010320"
      wRes = 0
      wCls = 1
      Call GroupCalculation(wDate, wDate, wDbt, wRes, wOut, wInc, wCls, wRsk)
      
      BuiltIn.Delay(1000)
      Call Close_Pttel("frmPttel")
      BuiltIn.Delay(4000)

      ' Մուտք Հաճախորդներ թղթապանակ 
      Call wTreeView.DblClickItem("|¸»µÇïáñ³Ï³Ý å³ñïù»ñ|ä³ÛÙ³Ý³·ñ»ñ")
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error(folderName & " դիալոգը չի բացվել")
      End If
      
      cliCode =  "8101800"
      ' Հաճախորդ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACCBAL", cliCode)
      ' Ցույց տալ փակվածները
      Call Rekvizit_Fill("Dialog", 2, "CheckBox", "CLOSE", 1)
      
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(5000)

       If  tdbgView.ApproxCount <> 1 Then
             Log.Error(" Դեբիտորական պարտք(Ետհաշվեկշռային) պայմանագրիը առկա չէ Պայմանագրեր թղթապանակում")
             Exit Sub
       End If

      dateCl = wMDIClient.VBObject("frmPttel").GetColumnIndex("fDATECLOSE")
      If  Not Trim( tdbgView.Columns.Item(dateCl).Value) = Trim("01/03/20")  Then
            Log.Error("Դեբիտորական պարտք(Ետհաշվեկշռային) պայմանագիրը չի փակվել")
      End If
      
      ' Պայմանագրի բացում
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_AgrOpen)
      Call ClickCmdButton(5, "²Ûá")
      
      If  Not Trim( tdbgView.Columns.Item(dateCl).Value) = Trim("") Then
            Log.Error("Պայմանագիրը չի փակվել")
      End If
      
      ' Կատարած գործողությունների ջնջում
      sDatePar = "START"
      eDatePar = "END"
      action = c_Delete
      dateGive = "010120"
      param = c_OpersView
      state = True
      Call DeleteFromAllActionDoc(param, sDatePar, dateGive, eDatePar, wDate, state, action)
      
      Call Close_Pttel("frmPttel_2")
      BuiltIn.Delay(2000)
      
      ' Ռիսկի դասիչ և պահուստավորման փաստաթղթերի ջնջում
      param = c_ViewEdit & "|" & c_Risking & "|" & c_RisksPersRes
      Call DeleteFromAllActionDoc(param, sDatePar, dateGive, eDatePar, wDate, state, action)
      
      Call Close_Pttel("frmPttel_2")
      BuiltIn.Delay(2000)

      ' Ջնջել պայմանագիրը                
      Call DelDoc()
      Call Close_Pttel("frmPttel")
      
      ' Փակել ՀԾ-Բանկ ծրագիրը
      Call Close_AsBank()        
      
End Sub