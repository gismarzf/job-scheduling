VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDiagramaShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private pShape As Shape
Private pColor As Long
'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************

Property Get Color() As Long
    Color = pColor
End Property

Property Let Color(l As Long)
    pColor = l
End Property

Property Get Shape() As Shape
    Set Shape = pShape
End Property

Property Set Shape(s As Shape)
    Set pShape = s
End Property

'********************************************************************************
'FUNCIONES
'********************************************************************************

Public Sub visible()
    Shape.visible = msoCTrue
End Sub

Public Sub invisible()
    Shape.visible = msoFalse
End Sub

Public Sub mandarAlFrente()
    Shape.ZOrder msoBringToFront
End Sub

Public Sub mandarAtras()
    Shape.ZOrder msoSendToBack
End Sub

Public Sub marcar()

    With Shape.Line
        .Style = msoLineThickBetweenThin
        .visible = msoTrue
        .DashStyle = msoLineSysDot
        .visible = msoTrue
        .Weight = weightLineMark
        .ForeColor.RGB = glowColor
        .Transparency = 0
    End With

End Sub

Public Sub marcarRemover()

    Shape.Line.Style = msoLineSingle
    With Shape.Line
        .Style = msoLineSingle
        .visible = msoTrue
        .Weight = weightLineNormal
        .DashStyle = msoLineSolid
        .ForeColor.RGB = Color
    End With
        
End Sub


