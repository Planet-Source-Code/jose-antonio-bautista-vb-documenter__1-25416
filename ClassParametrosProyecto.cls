VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassParProyecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase para guardar otros par�metros del proyecto
Option Explicit

Private Type tParametro 'Tipo para almacenar otros par�metros
  Tipo As String 'Tipo
  Descripcion As String 'Descripci�n
End Type

Private TablaParametros() As tParametro 'Tabla con los par�metros del proyecto

Public Sub Clear()
'--> Limpia la tabla
  ReDim TablaParametros(0)
End Sub

Public Sub Add(ByVal Identificador As String, ByVal Cadena As String)
'--> A�ade un elemento a la tabla de par�metros
  With TablaParametros(UBound(TablaParametros))
    .Tipo = Identificador
    If Left$(Cadena, 1) = Chr$(34) Then
      .Descripcion = Mid$(Cadena, 2, Len(Cadena) - 2)
    Else
      .Descripcion = Cadena
    End If
  End With
  ReDim Preserve TablaParametros(UBound(TablaParametros) + 1)
End Sub

Public Sub writeHTML(ByRef clGeneradorHTML As ClassHTML, ByVal NumeroFichero As Long)
'--> Escribe el HTML con los otros par�metros
Dim Indice As Integer
Dim CadSalida As String

  CadSalida = "<TR> <TD CLASS=TableSubcaption> Tipo"
  CadSalida = CadSalida + "<TD CLASS=TableSubcaption> Descripci�n"
  For Indice = 0 To UBound(TablaParametros) - 1
    With TablaParametros(Indice)
      CadSalida = CadSalida + "<TR>"
      CadSalida = CadSalida + "<TD>" + .Tipo
      CadSalida = CadSalida + "<TD>" + .Descripcion
    End With
  Next Indice
  Print #NumeroFichero, clGeneradorHTML.getTable("Par�metros del proyecto", CadSalida)
End Sub

Private Sub Class_Initialize()
  ReDim TablaParametros(0)
End Sub

Private Sub Class_Terminate()
  Erase TablaParametros
End Sub
