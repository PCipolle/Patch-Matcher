Attribute VB_Name = "Main"
Global numPatches As Integer
Global filePath As String
Global patchSizes(70, 1) As Variant


Public Sub patchMatcher(flag As Integer)
Dim frm As Frame
Set frm = App.Frame
Dim location As String
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim geos As Paths

location = frm.PathOfThisAddin

If flag = 1 Then

    numPatches = 70
    filePath = location & "/patchSizesTop.txt"
    initPatchSizes filePath
    Set geos = drw.UserSelectMultiGeosCollection("Select Patch Geometries", 0)
    searchAndReplace geos

    
ElseIf flag = 2 Then

    numPatches = 60
    filePath = location & "/patchSizesBottom.txt"
    initPatchSizes filePath
    Set geos = drw.UserSelectMultiGeosCollection("Select Patch Geometries", 0)
    searchAndReplace geos
    
ElseIf flag = 3 Then
    
    numPatches = 7
    filePath = location & "/woodButterflies.txt"
    initPatchSizes filePath
    Set geos = drw.UserSelectMultiGeosCollection("Select Patch Geometries", 0)
    butterfly geos

ElseIf flag = 4 Then
    
    numPatches = 2
    filePath = location & "/bronzeButterflies.txt"
    initPatchSizes filePath
    Set geos = drw.UserSelectMultiGeosCollection("Select Patch Geometries", 0)
    butterfly geos
    
End If

      
End Sub


