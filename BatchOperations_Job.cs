'USEUNIT Library_Common 
'USEUNIT Library_CheckDB 
'USEUNIT Payment_Library 
                                                  
Sub BatchOperations_Job_Test
'  Ÿµ≥€«› ∑·ÒÕ·’·ı√€·ı››ªÒ« ÔªÎÔ 


  Log.Message("BatchOperations_Job_Test Started")

  Utilities.ShortDateFormat = "yyyymmdd" 
  fDATE = Utilities.DateToStr(Utilities.Date())
  startDATE = Utilities.DateToStr(Utilities.IncMonth(Utilities.Now, -12)) 
    
  Call Initialize_AsBank("bank", startDATE, fDATE) 
  BuiltIn.Delay(1000)  
 
  Call wTreeView.DblClickItem("|≤πŸ«›«ÎÔÒ≥Ô·Ò« ≤ﬁÓ|≤È≥Á≥πÒ≥›˘›ªÒ|≤È≥Á≥πÒ≥›˘›ªÒ")
  Call p1.VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton  
   
  Set wfrmPttel =  wMDIClient.VBObject("frmPttel")
  wfrmPttel.SetFocus()
   
  For i=1 to wfrmPttel.VBObject("tdbgView").VisibleRows
    Call wMDIClient.VBObject("frmPttel").ClickR()
    BuiltIn.Delay(1000)
    Call wMainForm.PopupMenu.Click("Ê›Áª…")
    BuiltIn.Delay(1000)
    Call p1.VBObject("frmAsMsgBox").VBObject("cmdButton").ClickButton
  Next
 
  Call wMDIClient.VBObject("frmPttel").ClickR()
  BuiltIn.Delay(1000)
  Call wMainForm.PopupMenu.Click("≤Ìª…≥Û›ª…")
  BuiltIn.Delay(1000)
  Call p1.VBObject("frmEditJob").VBObject("ASGroupTree").VBObject("TDBMask").Keys("Testing" & "[Tab]")  
  Call p1.VBObject("frmEditJob").VBObject("RunButtom").ClickButton
  BuiltIn.Delay(90000) 
  Call p1.VBObject("frmJobSetExec").Close()
  'BuiltIn.Delay(5000)  
  Dim count
  count = 0   
  Do While count < 50
     BuiltIn.Delay(10000)
     If wfrmPttel.VBObject("tdbgView").VisibleRows = 1  Then  Exit Do
     count = count+1 
  Loop
        
  If Not wfrmPttel.VBObject("tdbgView").VisibleRows = 1  Then
    Log.Error("TimeOut!") 
    Exit Sub      
  End If
  
  If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").columns(2).text = "Œ≥Ô≥ÒÌ≥Õ ø" Then
    Log.Message("Job Completed Successfully")
  Else
    Log.Error("Job Completed Unsuccessfully")  
  End If
  
  Call Close_AsBank

  Log.Message("BatchOperations_Job_Test Completed")
  BuiltIn.Delay(3000)       
 
End Sub 