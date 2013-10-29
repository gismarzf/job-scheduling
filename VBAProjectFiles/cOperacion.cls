VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private pMaquina As cMaquina
Private pTrabajo As cTrabajo
Private pDuracion As Integer
Private pInicio As Integer
Private pFin As Integer
Private pCirculo As cDiagramaCirculo
Private pOrdenEnElTrabajo As Integer
Private pIndice As Integer
Private pArcos As New Collection
Private pHorarioInicio As Integer
Private pHorarioFin As Integer
Private pCompletado As Boolean
Private pDiagrama As cDiagrama
Private pRectangulo As cDiagramaRectangulo

'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************

Property Get Diagrama() As cDiagrama
    Set Diagrama = pDiagrama
End Property

Property Set Diagrama(d As cDiagrama)
    Set pDiagrama = d
End Property

Property Get Completado() As Boolean
    Completado = pCompletado
End Property

Property Let Completado(b As Boolean)
    pCompletado = b
End Property

Property Get HorarioInicio() As Integer
    HorarioInicio = pHorarioInicio
End Property

Property Let HorarioInicio(i As Integer)
    pHorarioInicio = i
End Property

Property Get HorarioFin() As Integer
    HorarioFin = pHorarioFin
End Property

Property Let HorarioFin(i As Integer)
    pHorarioFin = i
End Property

Property Get Arcos() As Collection
    Set Arcos = pArcos
End Property

Property Set Arcos(c As Collection)
    Set pArcos = c
End Property

Property Get indice() As Integer
    indice = pIndice
End Property

Property Let indice(i As Integer)
    pIndice = i
End Property

Property Get OrdenEnElTrabajo() As Integer
    OrdenEnElTrabajo = pOrdenEnElTrabajo
End Property

Property Let OrdenEnElTrabajo(i As Integer)
    pOrdenEnElTrabajo = i
End Property

Property Get Maquina() As cMaquina
    Set Maquina = pMaquina
End Property

Property Set Maquina(m As cMaquina)
    Set pMaquina = m
End Property

Property Get Trabajo() As cTrabajo
    Set Trabajo = pTrabajo
End Property

Property Set Trabajo(t As cTrabajo)
    Set pTrabajo = t
End Property

Property Get Duracion() As Integer
    Duracion = pDuracion
End Property

Property Let Duracion(i As Integer)
    pDuracion = i
End Property

Property Get Circulo() As cDiagramaCirculo
    Set Circulo = pCirculo
End Property

Property Set Circulo(c As cDiagramaCirculo)
    Set pCirculo = c
End Property

Property Get Rectangulo() As cDiagramaRectangulo
    Set Rectangulo = pRectangulo
End Property

Property Set Rectangulo(r As cDiagramaRectangulo)
    Set pRectangulo = r
End Property

'********************************************************************************
'FUNCIONES
'********************************************************************************


Public Function esPosibleHacer() As Boolean
    Dim b As Boolean
    b = True
    
    Dim a As cArco
    For Each a In Arcos
        ' busco la operacion de un arco cuya flecha me llega a mi (final soy yo) y veo si no fue completado todavia
        If a.FinalOperacion.indice = indice And a.InicioOperacion.Completado = False Then b = False
    Next
    
    esPosibleHacer = b
    
End Function

Public Function posiblesSucedores() As Collection
    
    Dim c As New Collection
    Dim a As cArco
    For Each a In Arcos
        If a.InicioOperacion.indice = indice Then c.Add a.FinalOperacion
    Next
    
    Set posiblesSucedores = c
    
End Function

Public Function posiblesPrecedores() As Collection
    
    Dim c As New Collection
    Dim a As cArco
    For Each a In Arcos
        If a.FinalOperacion.indice = indice Then c.Add a.InicioOperacion
    Next
    
    Set posiblesPrecedores = c
    
End Function

