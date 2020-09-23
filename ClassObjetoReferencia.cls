VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassObjetoReferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase de manejo de objetos / referencias del proyecto
Option Explicit

Private Type tObjetoReferencia 'Tipo para objetos o referencias
  isObject As Boolean 'Indice si es un objeto o una referencia
  NumeroRegistro As String 'Guarda el número de registro
  Version As String 'Guarda la versión
  SubVersion As String 'Guarda la revisión
  NombreFichero As String 'Guarda el nombre de fichero
  Descripcion As String 'En su caso guarda la descripción
End Type

Private TablaObjetos() As tObjetoReferencia 'Tabla con todos los objetos del proyecto

Public Sub Clear()
'--> Limpia la tabla
  ReDim TablaObjetos(0)
End Sub

Public Sub Add(ByVal Identificador As String, ByVal Cadena As String)
'--> Añade un elemento a la tabla de objetos (TablaObjetos)
  With TablaObjetos(UBound(TablaObjetos))
    .isObject = (UCase$(Identificador) = "OBJECT")
    .NumeroRegistro = guObjGeneral.QuitarParametro(Cadena, "#")
    .Version = guObjGeneral.QuitarParametro(Cadena, "#")
    .SubVersion = guObjGeneral.QuitarParametro(Cadena, IIf(.isObject, ";", "#"))
    .NombreFichero = guObjGeneral.QuitarParametro(Cadena, "#")
    .Descripcion = Cadena
  End With
  ReDim Preserve TablaObjetos(UBound(TablaObjetos) + 1)
End Sub

Private Function getHTMLObject(ByRef clGeneradorHTML As ClassHTML, ByVal isObject As Boolean) As String
'--> Devuelve una tabla HTML con objetos o referencias
Dim Indice As Integer
Dim CadSalida As String

  CadSalida = "<TR>"
  CadSalida = CadSalida + "<TD CLASS=TableSubcaption> Descripción"
  CadSalida = CadSalida + "<TD CLASS=TableSubcaption> Versión"
  CadSalida = CadSalida + "<TD CLASS=TableSubcaption> Número Registro"
  For Indice = 0 To UBound(TablaObjetos) - 1
    With TablaObjetos(Indice)
      If .isObject = isObject Then
        CadSalida = CadSalida + "<TR>"
        If isObject Then
          CadSalida = CadSalida + "<TD>" + .NombreFichero
        Else
          CadSalida = CadSalida + "<TD>" + .Descripcion + " (" + .NombreFichero + ")"
        End If
        CadSalida = CadSalida + "<TD>" + .Version
        CadSalida = CadSalida + "<TD>" + .NumeroRegistro
      End If
    End With
  Next Indice
  getHTMLObject = clGeneradorHTML.getTable(IIf(isObject, "Objetos", "Referencias"), CadSalida)
End Function

Public Sub writeHTML(ByRef clGeneradorHTML As ClassHTML, ByVal NumeroFichero As Long)
'--> Escribe el HTML con los objetos y referencias
  Print #NumeroFichero, getHTMLObject(clGeneradorHTML, True) + _
                        getHTMLObject(clGeneradorHTML, False)
End Sub

Private Sub Class_Initialize()
  ReDim TablaObjetos(0)
End Sub

Private Sub Class_Terminate()
  Erase TablaObjetos
End Sub
