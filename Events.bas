Attribute VB_Name = "Events"
 Function InitAlphacamAddIn(AcamVersion As Long) As Integer
  Dim frm As Frame
  Set frm = App.Frame
  Dim btnId As Long
  Dim location As String
  
  location = frm.PathOfThisAddin
  
  btnId = frm.CreateButtonBar("Patch Matcher")
  
  frm.AddMenuItem2 "&Match Patches", "matchPatches", acamMenuNEW, "Patch and Butterfly"
  frm.AddButton btnId, location + "\PatchButterfly.png", frm.LastMenuCommandID
    
      

  InitAlphacamAddIn = 0

 End Function

Sub matchPatches()
    frmSelect.StartUpPosition = 0
    frmSelect.Top = 10
    frmSelect.Left = 10
    frmSelect.Show
    
End Sub
