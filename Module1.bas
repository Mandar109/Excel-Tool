Attribute VB_Name = "Module1"
Option Explicit

Const totalSteps As Integer = 100
Const totalLoadN As Double = 5000000 ' 500 tons
Const plateWidthPx As Double = 400
Const maxDeflectionRadius As Double = 300 ' Radius in pixels

Sub SimulateBucklingArc()
    Dim stepNum As Integer
    Dim loadFraction As Double
    Dim arcShape As Shape
    Dim deflectionRadius As Double
    
    ' Clean up any previous shapes
    Dim s As Shape
    For Each s In ActiveSheet.Shapes
        If s.Name Like "BucklingArc*" Then s.Delete
    Next s

    ' Add title
    Range("A1").value = "Steel Plate Buckling (Arc Simulation)"
    Range("A2").value = "500x400mm Plate | 80mm Thick | Load: 500 tons"
    
    ' Animate buckling with arc
    For stepNum = 1 To totalSteps
        loadFraction = stepNum / totalSteps
        deflectionRadius = 10 + loadFraction * maxDeflectionRadius
        
        ' Draw the arc
        Set arcShape = DrawBucklingArc(deflectionRadius, stepNum)
        
        ' Show info
        Range("A4").value = "Step: " & stepNum & "/" & totalSteps
        Range("A5").value = "Load: " & Round(loadFraction * totalLoadN / 1000, 0) & " kN"
        Range("A6").value = "Radius: " & Round(deflectionRadius, 0) & " px"

        DoEvents
        WaitMilliseconds 50
        
        ' Delete previous arc
        If stepNum > 1 Then
            On Error Resume Next
            ActiveSheet.Shapes("BucklingArc" & (stepNum - 1)).Delete
            On Error GoTo 0
        End If
    Next stepNum
End Sub

Function DrawBucklingArc(radius As Double, stepNum As Integer) As Shape
    Dim centerX As Double, centerY As Double
    Dim arcWidth As Double, arcHeight As Double
    Dim arcLeft As Double, arcTop As Double
    Dim arcAngle As Double

    centerX = 100 + plateWidthPx / 2
    centerY = 250
    
    arcWidth = plateWidthPx
    arcHeight = radius * 2 ' Diameter is twice the radius

    arcLeft = centerX - arcWidth / 2
    arcTop = centerY - radius

    ' Draw symmetrical top arc (upward deflection)
    Set DrawBucklingArc = ActiveSheet.Shapes.AddShape(msoShapeArc, arcLeft, arcTop, arcWidth, arcHeight)
    With DrawBucklingArc
        .Name = "BucklingArc" & stepNum
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Line.Weight = 3
        
        ' Symmetrical arc: left to right 180 degrees
        .Adjustments.Item(1) = 0   ' Start angle
        .Adjustments.Item(2) = 180 ' End angle
    End With
End Function

Sub WaitMilliseconds(ms As Long)
    Dim endTime As Double
    endTime = Timer + ms / 1000#
    Do While Timer < endTime
        DoEvents
    Loop
End Sub

