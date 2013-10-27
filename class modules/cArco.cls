VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cArco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private pInicioOperacion As cOperacion
Private pFinalOperacion As cOperacion
Private pConector As cDiagramaConector
Private pDireccion As Boolean
Private pMaquina As cMaquina
Private pTrabajo As cTrabajo
Private pTipo As String
Private pDiagrama As cDiagrama
Private pIndice As Integer

'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************

Property Get indice() As Integer
    indice = pIndice
End Property

Property Let indice(i As Integer)
    pIndice = i
End Property

Property Get Diagrama() As cDiagrama
    Set Diagrama = pDiagrama
End Property

Property Set Diagrama(d As cDiagrama)
    Set pDiagrama = d
End Property

Property Get Conector() As cDiagramaConector
    Set Conector = pConector
End Property

Property Set Conector(d As cDiagramaConector)
    Set pConector = d
End Property

Property Get Tipo() As String
    Tipo = pTipo
End Property

Property Let Tipo(s As String)
    pTipo = s
End Property

' solo si es arco disyuntivo
Property Get Maquina() As cMaquina
    Set Maquina = pMaquina
End Property

Property Set Maquina(m As cMaquina)
    Set pMaquina = m
End Property

' solo si es arco conyuntivo
Property Get Trabajo() As cTrabajo
    Set Trabajo = pTrabajo
End Property

Property Set Trabajo(t As cTrabajo)
    Set pTrabajo = t
End Property

Property Get Direccion() As Boolean
    Direccion = pDireccion
End Property

Property Let Direccion(b As Boolean)
    pDireccion = b
End Property

Property Get InicioOperacion() As cOperacion

    ' si la direccion es uno, la flecha va de una operacion con indice menor a otra
    ' con indice mayor, sino al reves
    
    If Direccion = True And pInicioOperacion.indice < pFinalOperacion.indice Then
        Set InicioOperacion = pInicioOperacion
    ElseIf Direccion = True And pInicioOperacion.indice > pFinalOperacion.indice Then
        Set InicioOperacion = pFinalOperacion
    ElseIf Direccion = False And pInicioOperacion.indice < pFinalOperacion.indice Then
        Set InicioOperacion = pFinalOperacion
    ElseIf Direccion = False And pInicioOperacion.indice > pFinalOperacion.indice Then
        Set InicioOperacion = pInicioOperacion
    End If
    
End Property

Property Set InicioOperacion(c As cOperacion)
    Set pInicioOperacion = c
End Property

Property Get FinalOperacion() As cOperacion
    
    ' si la direccion es uno, la flecha va de una operacion con indice menor a otra
    ' con indice mayor, sino al reves
    
    If Direccion = True And pInicioOperacion.indice < pFinalOperacion.indice Then
        Set FinalOperacion = pFinalOperacion
    ElseIf Direccion = True And pInicioOperacion.indice > pFinalOperacion.indice Then
        Set FinalOperacion = pInicioOperacion
    ElseIf Direccion = False And pInicioOperacion.indice < pFinalOperacion.indice Then
        Set FinalOperacion = pInicioOperacion
    ElseIf Direccion = False And pInicioOperacion.indice > pFinalOperacion.indice Then
        Set FinalOperacion = pFinalOperacion
    End If
    
End Property

Property Set FinalOperacion(c As cOperacion)
    Set pFinalOperacion = c
End Property

'********************************************************************************
'FUNCIONES
'********************************************************************************

Public Sub agregarConectorADiagrama()


End Sub
