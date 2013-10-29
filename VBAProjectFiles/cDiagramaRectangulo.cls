VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDiagramaRectangulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pmiShape As cDiagramaShape
Private pX As Integer, pY As Integer
Private pWidth As Integer, pHeight As Integer
Private pOperacion As cOperacion

'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************

Property Get operacion() As cOperacion
    Set operacion = pOperacion
End Property

Property Set operacion(o As cOperacion)
    Set pOperacion = o
End Property

Property Get Width() As Integer
    Width = pWidth
End Property

Property Let Width(i As Integer)
    pWidth = i
End Property

Property Get Height() As Integer
    Height = pHeight
End Property

Property Let Height(i As Integer)
    pHeight = i
End Property

Property Get x() As Integer
    x = pX
End Property

Property Let x(i As Integer)
    pX = i
End Property

Property Get y() As Integer
    y = pY
End Property

Property Let y(i As Integer)
    pY = i
End Property

Property Get miShape() As cDiagramaShape
    Set miShape = pmiShape
End Property

Property Set miShape(s As cDiagramaShape)
    Set pmiShape = s
End Property
