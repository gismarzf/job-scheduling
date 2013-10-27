Attribute VB_Name = "Helper"
Option Explicit
Option Base 1

'********************************************************************************
'CONSTANTES
'********************************************************************************

' diag disyuntivo
Public Const widthCirculo As Integer = 40
Public Const heightCirculo As Integer = 40
Public Const zeroXCirculo As Integer = 10
Public Const zeroYCirculo As Integer = 10
Public Const separacionVerticalCirculos As Integer = 100
Public Const weightLineNormal As Integer = 2
Public Const weightLineHidden As Integer = 1
Public Const weightLineMark As Integer = 2

' diag gantt
Public Const zeroYRectangulo As Integer = 2
Public Const heightRectangulo As Integer = 30
Public Const factorWidthRectangulo As Double = 0.6
Public Const separacionVerticalRectangulos As Integer = 5
Public Const numLineasSepaGantt As Integer = 10
Public Const distanciaTop As Integer = 25

'colores
Public Const minColor As Long = 3289650 + 10000
Public Const maxColor As Long = 13487565 - 10000
Public Const glowColor As Long = 255
Public Const colorDif As Double = 0.8

'********************************************************************************
' VARIABLES
'********************************************************************************
' para benchmark
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public start As Long, finish As Long ' benchmark

Public MaxListaTabu As Integer, maxListaTabuInicial As Integer ' segundo es para la sol. inicial

' variables para configurar las planillas distintas
Public separacionHorizontalCirculos As Integer, zeroXRectangulo As Integer
Public Factor_Gantt As Double

'********************************************************************************
'FUNCIONES
'********************************************************************************

Public Function coleccionContieneObjeto(item As Variant, c As Collection)
    
    If c Is Nothing Then
        coleccionContieneObjeto = False
    Else
        Dim v As Variant
        For Each v In c
            If item Is v Then coleccionContieneObjeto = True
        Next
    End If
    
End Function

Public Function coleccionContieneInteger(valor As Integer, c As Collection) As Boolean
    
    If c Is Nothing Then
        coleccionContieneInteger = False
    Else
        Dim i As Integer
        For i = 1 To c.count
            If c(i) = valor Then coleccionContieneInteger = True
        Next i
    End If

End Function

' supongo que es mas rapido ya que no crea una nueva coleccion
Public Sub agregarColeccion(c1 As Collection, c2 As Collection)
    Dim v As Variant
    For Each v In c2
        c1.Add v
    Next
End Sub

Public Function sumarColecciones(c1 As Collection, c2 As Collection) As Collection
        
    Dim csum As New Collection
    
    Dim v As Variant
    For Each v In c1
        csum.Add v
    Next

    For Each v In c2
        csum.Add v
    Next

    Set sumarColecciones = csum

    Set csum = Nothing

End Function

Public Function performanceOn()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
End Function

Public Function performanceOff()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Function

Public Function vectorReemplazoPrimerCero(i As Integer, v() As Integer) As Integer()
    
    Dim k As Integer
    For k = 1 To UBound(v, 1)
        If v(k) = 0 Then
            v(k) = i
            vectorReemplazoPrimerCero = v
            Exit Function
        End If
    Next k
    
End Function

Public Function vectorContieneInteger(i As Integer, v() As Integer) As Boolean
    Dim k As Integer
    For k = 1 To UBound(v, 1)
        If v(k) = i Then
            vectorContieneInteger = True
            Exit Function
        End If
    Next k
End Function

