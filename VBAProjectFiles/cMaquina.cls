VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMaquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private pNombre As String
Private pNumero As Integer
Private pColor As Long

'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************

Property Get Nombre() As String
    Nombre = pNombre
End Property

Property Let Nombre(s As String)
    pNombre = s
End Property

Property Get Numero() As Integer
    Numero = pNumero
End Property

Property Let Numero(i As Integer)
    pNumero = i
End Property

Property Get Color() As Long
    Color = pColor
End Property

Property Let Color(l As Long)
    pColor = l
End Property

