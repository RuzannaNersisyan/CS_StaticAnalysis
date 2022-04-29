'USEUNIT Library_Common 
'USEUNIT Library_CheckDB 
'USEUNIT Payment_Library 

Sub BatchOperations_Test
'  Ÿµ≥€«› ∑·ÒÕ·’·ı√€·ı››ªÒ« ÔªÎÔ


  Log.Message("BatchOperations_Test Started")
  
  Call Initialize_AsBank("bank", "20060505", "20090101")
  
  Login ("creditoperator")
  BuiltIn.Delay(3000)
  Call wTreeView.DblClickItem("|Ï≥ÒœªÒ (Ôª’≥µ≥ﬂÀÌ≥Õ)| Ÿµ≥€«› ∑·ÒÕ·’·ı√€·ı››ªÒ")
  BuiltIn.Delay(3000)
  Set wfrmAsUstPar = p1.VBObject("frmAsUstPar")
  BuiltIn.Delay(1000)
  Set wTabFrame = wfrmAsUstPar.VBObject("TabFrame")
  BuiltIn.Delay(500)
  Call wTabFrame.VBObject("Checkbox").ClickButton(cbUnChecked)
  BuiltIn.Delay(500) 
  Call wTabFrame.VBObject("Checkbox_2").ClickButton(cbChecked)
  BuiltIn.Delay(500) 
  Call wTabFrame.VBObject("Checkbox_8").ClickButton(cbChecked)
  BuiltIn.Delay(500) 
  Call wTabFrame.VBObject("Checkbox_9").ClickButton(cbChecked)
  BuiltIn.Delay(500) 
  Call wTabFrame.VBObject("Checkbox_10").ClickButton(cbChecked)
  BuiltIn.Delay(500) 
  Call wfrmAsUstPar.VBObject("CmdOK").ClickButton
  BuiltIn.Delay(500)  
     
  Set wFrmSpr = wMDIClient.WaitVBObject("FrmSpr", 300000)   
  
  If wFrmSpr.Exists Then
    Call wMDIClient.VBObject("FrmSpr").SetFocus()
    BuiltIn.Delay(3000)       
    Call wMDIClient.VBObject("FrmSpr").Close()
    BuiltIn.Delay(3000)  
  Else
    Call Log.Error("FrmSpr object not found")
  End If
  BuiltIn.Delay(3000)
  Call Close_AsBank

  Log.Message("BatchOperations_Test Completed")
    
  BuiltIn.Delay(3000)       
  
End Sub 

