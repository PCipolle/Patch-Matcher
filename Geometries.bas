Attribute VB_Name = "Geometries"
Public Function getLength(geo As Path)

Dim length As Double
Dim tempLength As Double
Dim realLength As Double
Dim testPoint As Double
Dim index As Integer

realLength = geo.length
length = 100

            For i = 0 To numPatches - 1
                tempLength = patchSizes(i, 0)
                If Abs(realLength - tempLength) < Abs(realLength - length) Then
                    length = tempLength
                    index = i
                End If
                
            Next i
 getLength = patchSizes(index, 1)

End Function

Public Function getAngle(geo As Path)
Dim x1 As Double
Dim x2 As Double
Dim y1 As Double
Dim y2 As Double
Dim angle As Double
Dim result As Boolean


x1 = geo.MinXL
x2 = geo.MaxXL
y1 = geo.MinYL
y2 = geo.MaxYL

result = geo.IsPointInside(geo.MinXL, geo.MaxYL)
If (x2 - x1) = 0 Then
    angle = 90
ElseIf (y2 - y1) = 0 Then
    angle = 0
Else
    angle = Math.Atn((y2 - y1) / (x2 - x1))
    angle = angle * 180 / 3.14159265359
End If

If result = False Then
    getAngle = 90 - angle
ElseIf result = True Then
    getAngle = angle - 90
End If

End Function

Public Function butterfly(geos As Paths)
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim line1 As Double
Dim line2 As Double
Dim angle As Double
Dim offset As Double
Dim diam1 As Double
Dim diam2 As Double
Dim circOffset As Double
Dim zBottom As Double
Dim lineAngle As Double
Dim geo As Path
Dim geo1 As Path
Dim geo2 As Path
Dim geo3 As Path
Dim geo4 As Path
Dim geo5 As Path
Dim geo6 As Path
Dim circ As Path
Dim flag As String
Dim xMid As Double
Dim yMid As Double
Dim butterflyGeos As Paths
Dim butterflyGeo As Path
Dim radius1 As Double
Dim radius2 As Double
Dim zAdjust As Double
Dim butterflies As Paths
Dim toolPathFlag As Integer
Dim cornerCircle As Path
Dim cornerCircles As Paths
Dim circles As Paths


toolPathFlag = MsgBox("Insert Tool Paths?", vbYesNo, "Tool Paths")

Set butterflies = drw.CreatePathCollection
Set circles = drw.CreatePathCollection
Set cornerCircles = drw.CreatePathCollection

zAdjust = InputBox("Enter Z Level Adjustment:", "Z Adjust", 0, 10, 10)

For Each geo In geos

    flag = getLength(geo)
    
    lineAngle = getAngle(geo)
    
    If flag = "W1" Then
    
        diam1 = 0.05
        diam2 = 0.05
        circOffset = 0.35
        line1 = 0.63279
        offset = 0.20458
        angle = 9
        zBottom = -0.5
        radius1 = 0
        radius2 = 0
        
    ElseIf flag = "W2" Then
    
        diam1 = 0.05
        diam2 = 0.05
        circOffset = 0.35
        line1 = 0.75935
        offset = 0.26708
        angle = 9
        zBottom = -0.5
        radius1 = 0
        radius2 = 0
        
    ElseIf flag = "W3" Then
    
        diam1 = 0.05
        diam2 = 0.05
        circOffset = 0.6
        line1 = 1.01247
        offset = 0.30823
        angle = 9
        zBottom = -0.5
        radius1 = 0
        radius2 = 0
        
    ElseIf flag = "W4" Then
    
        diam1 = 0.05
        diam2 = 0.05
        circOffset = 0.6
        line1 = 1.13902
        offset = 0.39364
        angle = 9
        zBottom = -0.5
        radius1 = 0
        radius2 = 0
        
    ElseIf flag = "W5" Then
    
        diam1 = 0.05
        diam2 = 0.05
        circOffset = 0.975
        line1 = 1.5187
        offset = 0.39985
        angle = 9
        zBottom = -0.5
        radius1 = 0
        radius2 = 0
        
    ElseIf flag = "W6" Then
    
        diam1 = 0.05
        diam2 = 0.05
        circOffset = 1.475
        line1 = 2.02493
        offset = 0.49146
        angle = 9
        zBottom = -0.5
        radius1 = 0
        radius2 = 0
        
    ElseIf flag = "W7" Then
    
        diam1 = 0.05
        diam2 = 0.05
        circOffset = 2.118
        line1 = 2.66699
        offset = 0.47879
        angle = 7.77074
        zBottom = -0.5
        radius1 = 0
        radius2 = 0
        
    ElseIf flag = "B1" Then
    
        diam1 = 0.05
        diam2 = 0.25
        circOffset = 1
        line1 = 1.39837
        offset = 0.36335
        angle = 7.944815
        zBottom = -0.7
        radius1 = 0
        radius2 = 0.368
        
    ElseIf flag = "B2" Then
    
        diam1 = 0.05
        diam2 = 0.25
        circOffset = 1.374
        line1 = 1.89837
        offset = 0.38002
        angle = 8.25281
        zBottom = -0.7
        radius1 = 0
        radius2 = 0.362
    
    End If
    
    Set geo1 = drw.Create2DLine(-offset / 2, 0, -offset / 2, line1)
    Set geo2 = drw.Create2DLine(offset / 2, 0, offset / 2, line1)
    Set geo3 = drw.Create2DLine(-offset / 2, 0, -offset / 2, -line1)
    Set geo4 = drw.Create2DLine(offset / 2, 0, offset / 2, -line1)
    
    geo1.RotateL angle, -offset / 2, 0
    geo2.RotateL -angle, offset / 2, 0
    geo3.RotateL -angle, -offset / 2, 0
    geo4.RotateL angle, offset / 2, 0
        
    Set geo5 = drw.Create2DLine(geo1.MinXL, geo1.MaxYL, geo2.MaxXL, geo2.MaxYL)
    Set geo6 = drw.Create2DLine(geo3.MinXL, geo3.MinYL, geo4.MaxXL, geo4.MinYL)
                       
    geo1.ToolSide = acamRIGHT
    geo1.Reverse
    geo2.ToolSide = acamLEFT
    geo3.ToolSide = acamLEFT
    geo4.ToolSide = acamRIGHT
    geo4.Reverse
    geo5.ToolSide = acamRIGHT
    geo5.Reverse
    geo6.ToolSide = acamLEFT
    
    xMid = (geo1.MaxXL + geo1.MinXL) / 2
    yMid = (geo1.MaxYL + geo1.MinYL) / 2
    
    geo1.AddPath geo3
    geo1.Fillet radius2
    geo4.AddPath geo2
    geo4.Fillet radius2
    geo5.AddPath geo1
    geo5.AddPath geo4
    geo5.ToolSide = acamRIGHT
    geo5.Fillet radius1
    geo6.Delete
         
    geo1.AddPath geo5
    geo5.Fillet radius1
    
    geo5.SetStartPoint xMid, yMid
    geo5.ToolSide = acamLEFT

    geo5.Attribute("LicomUKDMBGeoZLevelTop") = 0#
    geo5.Attribute("LicomUKDMBGeoZLevelBottom") = zBottom + zAdjust
    
    xMid = (geo.MaxXL + geo.MinXL) / 2
    yMid = (geo.MaxYL + geo.MinYL) / 2
    
    Set cornerCircle = drw.CreateCircle(0.05, geo5.MinXL + 0.033, geo5.MinYL + 0.028)
    cornerCircle.RotateL lineAngle, 0, 0
    cornerCircle.MoveL xMid, yMid
    cornerCircles.Add cornerCircle
    Set cornerCircle = drw.CreateCircle(0.05, geo5.MinXL + 0.033, geo5.MaxYL - 0.028)
    cornerCircle.RotateL lineAngle, 0, 0
    cornerCircle.MoveL xMid, yMid
    cornerCircles.Add cornerCircle
    Set cornerCircle = drw.CreateCircle(0.05, geo5.MaxXL - 0.033, geo5.MaxYL - 0.028)
    cornerCircle.RotateL lineAngle, 0, 0
    cornerCircle.MoveL xMid, yMid
    cornerCircles.Add cornerCircle
    Set cornerCircle = drw.CreateCircle(0.05, geo5.MaxXL - 0.033, geo5.MinYL + 0.028)
    cornerCircle.RotateL lineAngle, 0, 0
    cornerCircle.MoveL xMid, yMid
    cornerCircles.Add cornerCircle
    
    Set circ = drw.CreateCircle(diam1, 0, 0)
    circ.RotateL lineAngle, 0, 0
    circ.MoveL xMid, yMid
    circles.Add circ
    
    Set circ = drw.CreateCircle(diam2, 0, circOffset)
    circ.RotateL lineAngle, 0, 0
    circ.MoveL xMid, yMid
    If lineAngle <> 0 And lineAngle <> 90 Then
        circles.Add circ
    End If
    
    Set circ = drw.CreateCircle(diam2, 0, -circOffset)
    circ.RotateL lineAngle, 0, 0
    circ.MoveL xMid, yMid
    If lineAngle <> 0 And lineAngle <> 90 Then
        circles.Add circ
    End If
    
    geo5.RotateL lineAngle, 0, 0
    geo5.MoveL xMid, yMid
    
    butterflies.Add geo5
    
Next geo

If toolPathFlag = 6 Then
    
    holesToolPath circles, -0.15, 100
    holesToolPath cornerCircles, -0.2, 150
    butterflyToolPath butterflies
End If

drw.RedrawShadedViews
drw.Redraw
drw.Refresh

End Function

Public Function searchAndReplace(geos As Paths)
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim geo As Path
Dim xGeo As Double
Dim yGeo As Double
Dim xReal As Double
Dim yReal As Double
Dim xTemp As Double
Dim yTemp As Double
Dim i As Integer
Dim xIndex As Integer
Dim xCenter As Double
Dim yCenter As Double
Dim patches As Paths
Dim patch As Path
Dim x As Double
Dim y As Double
Dim index As Integer
Dim newPatch As Path
Dim xAdjust As Double
Dim yAdjust As Double
Dim circ As Path
Dim newPatches As Paths
Dim centerHoles As Paths
Dim toolPathFlag As Integer
Dim cornerCircle As Path
Dim cornerCircles As Paths

toolPathFlag = MsgBox("Insert Tool Paths?", vbYesNo, "Tool Paths")

Set newPatches = drw.CreatePathCollection
Set centerHoles = drw.CreatePathCollection
Set cornerCircles = drw.CreatePathCollection

xAdjust = 0
yAdjust = 0

xAdjust = InputBox("Enter X Adjustment:", "X Adjust", 0.002, 10, 10)
yAdjust = InputBox("Enter Y Adjustment:", "Y Adjust", 0, 10, 10)


If geos.Count > 0 Then
      
    For Each geo In geos
    
        xReal = 100
        yReal = 100
        xGeo = geo.GetRectangleProperties.Width
        yGeo = geo.GetRectangleProperties.Height
            
            For i = 0 To numPatches - 1
                xTemp = patchSizes(i, 0)
                If Abs(xGeo - xTemp) < Abs(xGeo - xReal) Then
                    xReal = xTemp
                    xIndex = i
                End If
                
            Next i
            
            For i = xIndex To numPatches - 1
                yTemp = patchSizes(i, 1)
                If Abs(yGeo - yTemp) < Abs(yGeo - yReal) Then
                    yReal = yTemp
                    index = i
                End If
            Next i
        
        Set newPatch = drw.CreateRectangle(0, 0, patchSizes(index, 0) + xAdjust, patchSizes(index, 1) + yAdjust)
        
        x = newPatch.GetRectangleProperties.Right - (newPatch.GetRectangleProperties.Width / 2)
        y = newPatch.GetRectangleProperties.Top - (newPatch.GetRectangleProperties.Height / 2)
               
        xCenter = geo.GetRectangleProperties.Right - (geo.GetRectangleProperties.Width / 2)
        yCenter = geo.GetRectangleProperties.Top - (geo.GetRectangleProperties.Height / 2)

        newPatch.Attribute("LicomUKDMBGeoZLevelTop") = 0#
        newPatch.Attribute("LicomUKDMBGeoZLevelBottom") = -0.5
        
        newPatch.MoveL xCenter - x, yCenter - y
                      
        Set circ = drw.CreateCircle(0.05, xCenter, yCenter)
        
        newPatch.ToolSide = acamLEFT
        newPatch.ToolInOut = acamINSIDE
        newPatch.CW = False
        newPatch.SetStartPoint newPatch.MaxXL, ((newPatch.MaxYL + newPatch.MinYL) / 2)
        
        Set cornerCircle = drw.CreateCircle(0.05, newPatch.MinXL + 0.028, newPatch.MinYL + 0.028)
        cornerCircles.Add cornerCircle
        Set cornerCircle = drw.CreateCircle(0.05, newPatch.MinXL + 0.028, newPatch.MaxYL - 0.028)
        cornerCircles.Add cornerCircle
        Set cornerCircle = drw.CreateCircle(0.05, newPatch.MaxXL - 0.028, newPatch.MaxYL - 0.028)
        cornerCircles.Add cornerCircle
        Set cornerCircle = drw.CreateCircle(0.05, newPatch.MaxXL - 0.028, newPatch.MinYL + 0.028)
        cornerCircles.Add cornerCircle
        
        
        newPatches.Add newPatch
        centerHoles.Add circ
        
    Next geo
    
End If

If toolPathFlag = 6 Then

    holesToolPath centerHoles, -0.15, 100
    holesToolPath cornerCircles, -0.2, 60
    interiorPocketingToolPath newPatches
    
End If


drw.RedrawShadedViews
  
geos.Delete

End Function




Public Function initPatchSizes(filePath As String)

Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim line As String
Dim dLine As Double
Dim frm As Frame
Set frm = App.Frame

Dim i As Integer
i = 0

Open filePath For Input As #1

For i = 0 To numPatches - 1

    Line Input #1, line
    patchSizes(i, 0) = line
    
    
    Line Input #1, line
    patchSizes(i, 1) = line
          
Next i

Close #1

End Function
