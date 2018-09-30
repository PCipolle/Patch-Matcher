Attribute VB_Name = "Operations"
Public Function interiorPocketingToolPath(geos As Paths)
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim md As MillData
Set md = App.CreateMillData
Dim frm As Frame
Set frm = App.Frame
Dim toolLocation As String
Dim ld As LeadData

toolLocation = frm.PathOfThisAddin & "/Tools/.75 ROUGHER.art"
App.SelectTool toolLocation
md.PocketType = acamPocketCONTOUR
md.AutoZ = True
md.FinalPassIsland = acamFinalPassFULL
md.StartCutting = acamStartINSIDE
md.PocketUsePreviousMachining = True
md.SafeRapidLevel = 1
md.FeedDownDistance = 0.25
md.AutoZMaxDepthOfCut = 0.5625
md.Stock = 0.02
md.WidthOfCut = 0.45
md.OverlapOnOpenElements = 1
md.AutoZRampAngle = 30
geos.Selected = True
md.Pocket

Set md = App.CreateMillData
Set ld = App.CreateLeadData

toolLocation = frm.PathOfThisAddin & "/Tools/DOWNSHEAR 0.125 Inch.art"
App.SelectTool toolLocation
md.AutoZ = True
md.RoughFinishUsePreviousMachining = True
md.XYCorners = acamCornersSTRAIGHT
md.SafeRapidLevel = 1
md.FeedDownDistance = 0.25
md.Stock = -0.001
ld.LeadIn = acamLeadBOTH
ld.LeadOut = acamLeadBOTH
ld.RadiusIn = 2
ld.RadiusOut = 2
md.SetLeadData ld

geos.Selected = True

md.RoughFinish

geos.Selected = False

End Function

Public Function butterflyToolPath(geos As Paths)
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim md As MillData
Set md = App.CreateMillData
Dim frm As Frame
Set frm = App.Frame
Dim toolLocation As String
Dim ld As LeadData

toolLocation = frm.PathOfThisAddin & "/Tools/DOWNSHEAR 0.25 Inch.art"
App.SelectTool toolLocation
md.PocketType = acamPocketCONTOUR
md.AutoZ = True
md.FinalPassIsland = acamFinalPassFULL
md.StartCutting = acamStartINSIDE
md.PocketUsePreviousMachining = True
md.SafeRapidLevel = 1
md.FeedDownDistance = 0.25
md.AutoZMaxDepthOfCut = 0.25
md.AutoZRampAngle = 30
md.Stock = 0.02
md.WidthOfCut = 0.125

geos.Selected = True
md.Pocket

Set md = App.CreateMillData
Set ld = App.CreateLeadData

toolLocation = frm.PathOfThisAddin & "/Tools/DOWNSHEAR 0.125 Inch.art"
App.SelectTool toolLocation
md.AutoZ = True
md.RoughFinishUsePreviousMachining = True
md.XYCorners = acamCornersSTRAIGHT
md.SafeRapidLevel = 1
md.FeedDownDistance = 0.25
md.DepthOfCut = 0.5
md.Stock = -0.003
ld.LeadIn = acamLeadBOTH
ld.LeadOut = acamLeadBOTH
ld.RadiusIn = 1
ld.RadiusOut = 1
md.SetLeadData ld

geos.Selected = True

md.RoughFinish

End Function


Public Function holesToolPath(geos As Paths, depth As Double, feed As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim md As MillData
Set md = App.CreateMillData
Dim frm As Frame
Set frm = App.Frame
Dim toolLocation As String
Dim td As MillTool
Dim ms As MillStyle

toolLocation = frm.PathOfThisAddin & "/Tools/0.05 DRILL MILL.art"
Set td = App.SelectTool(toolLocation)
td.FixedDownFeed = feed
md.DrillType = acamDRILL
md.SafeRapidLevel = 1
md.RapidDownTo = 0.25
md.BottomOfHole = depth

geos.Selected = True
md.DrillTap
geos.Selected = False

End Function
