VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase que almacena los controles de un formulario
Option Explicit

Private Type tControl 'Tipo para la tabla de controles
  Nombre As String 'Nombre del control
  Tipo As String 'Tipo del control
End Type

Private TablaControles() As tControl 'Tabla con los controles del formulario

Public Sub addControl(ByVal Cadena As String)
'--> Añade un control a la tabla
  guObjGeneral.QuitarParametro Cadena, " "
  With TablaControles(UBound(TablaControles))
    .Tipo = guObjGeneral.QuitarParametro(Cadena, " ")
    .Nombre = Cadena
  End With
  If UCase$(TablaControles(UBound(TablaControles)).Tipo) <> "VB.FORM" Then
    ReDim Preserve TablaControles(UBound(TablaControles) + 1)
  End If
End Sub

Public Sub writeHTML(ByRef clGeneradorHTML As ClassHTML, ByVal NumeroFichero As Long)
'--> Escribe los controles del formulario sobre el fichero HTML
Dim Indice As Integer
Dim CadSalida As String

  If UBound(TablaControles) - 1 > 0 Then
    CadSalida = "<TR>"
    CadSalida = CadSalida + "<TD CLASS=TableSubcaption> Nombre"
    CadSalida = CadSalida + "<TD CLASS=TableSubcaption> Tipo"
    For Indice = 0 To UBound(TablaControles) - 1
      CadSalida = CadSalida + "<TR>"
      CadSalida = CadSalida + "<TD>" + TablaControles(Indice).Nombre
      CadSalida = CadSalida + "<TD>" + TablaControles(Indice).Tipo
    Next Indice
    Print #NumeroFichero, clGeneradorHTML.getAnchor("Controles") + clGeneradorHTML.getTable("Controles", CadSalida, False)
  End If
End Sub

Private Sub Class_Initialize()
  ReDim TablaControles(0)
End Sub

Private Sub Class_Terminate()
  Erase TablaControles
End Sub
