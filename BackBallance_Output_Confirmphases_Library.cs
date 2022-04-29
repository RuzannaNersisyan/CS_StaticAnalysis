Option Explicit
'USEUNIT Library_Common
'USEUNIT Online_PaySys_Library
'USEUNIT Constants

'----------------------------------------------
'ºïÑ³ßí»ÏßéÇ »ÉùÇ ûñ¹»ñ ÷³ëï³ÃÕÃÇ Éñ³óáõÙ
'----------------------------------------------
'docNumber - ö³ëï³ÃÕÃÇ Ñ³Ù³ñÁ
'summa - ¶áõÙ³ñ ¹³ßïÇ ³ñÅ»ù
'nbAcc - Ð³ßÇí ¹³ßïÇ ³ñÅ»ù
'aim - Üå³ï³Ï ¹³ßïÇ ³ñÅ»ù
'fISN - ö³ëï³ïÃÕÃÇ ISN-Á
'draft - true ³ñÅ»ùÇ ¹»åùáõÙ ë»ÕÙíáõÙ ¿ êñ³·Çñ Ïá×³ÏÁ, false-Ç ¹»åùáõÙ` Î³ï³ñ»É
Sub BackBallance_Output_Doc_Fill(docNumber, nbAcc, summa, aim, fISN, draft)
  Dim rekvName
    
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click("Ետհաշվեկշիռ|Ելքի օրդեր")
  BuiltIn.Delay(1000)
  wMDIClient.Refresh
   
  'êï»ÕÍíáÕ ISN - Ç ÷³ëï³ïÃÕÃÇ  í»ñ³·ñáõÙ ÷á÷áË³Ï³ÝÇÝ
  fISN = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
  'ö³ëï³ÃÕÃÇ N ¹³ßïÇ ³ñÅ»ùÇ í»ñ³·ñáõÙ ÷á÷áË³Ï³ÝÇÝ
  rekvName = GetVBObject("NOMDOK", wMDIClient.vbObject("frmASDocForm"))
  docNumber = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(rekvName).Text
    
  'Ð³ßÇí ¹³ßïÇ Éñ³óáõÙ
  Call Rekvizit_Fill("Document", 1, "General", "NBACC", nbAcc)
  '¶áõÙ³ñ ¹³ßïÇ Éñ³óáõÙ
  Call Rekvizit_Fill("Document", 1, "General", "SUMMA", summa)
  'Üå³ï³Ï ¹³ßïÇ Éñ³óáõÙ
  Call Rekvizit_Fill("Document", 1, "General", "COM", aim)
    
  ' Î³ï³ñ»É Ï³Ù ê¨³·Çñ Ïá×³ÏÇ ë»ÕÙáõÙ
  If draft Then
    Call ClickCmdButton(1, "ê¨³·Çñ")
  Else
    Call ClickCmdButton(1, "Î³ï³ñ»É")
  End If
End Sub