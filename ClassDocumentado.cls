VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDocumentado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase con los parámetros de los proyectos generados por el documentador
Option Explicit

Public strName As String 'Descripción corta del proyecto
Public strDescription As String 'Descripción larga
Public strFileProject As String 'Nombre completo de fichero del proyecto
Public strFileIndex As String 'Nombre completo del fichero de índice de HTML
Public strDateGeneration As String 'Fecha de generación
Public blnWithParameters As Boolean 'Si genera o no información sobre los parámetros de las rutinas
Public blnWithVariables As Boolean 'Si genera o no información sobre las variables locales de las rutinas

Private Type tDocRelacionado 'Tipo para el almacenaje de los documentos relacionados
  strDescription As String 'Descripción del documento relacionado
  strFileName As String 'Nombre del documento relacionado
End Type

Private uarrDocRelacionado() As tDocRelacionado 'Array con los documentos relacionados con la documentación

Public Sub Clear()
'--> Limpia la tabla de documentos relacionados
  ReDim uarrDocRelacionado(0)
End Sub

Public Sub addItem(ByVal strNewDescription As String, ByVal strNewFileName As String)
'--> Añade un elemento a los documentos relacionados
  With uarrDocRelacionado(UBound(uarrDocRelacionado))
    .strDescription = strNewDescription
    .strFileName = strNewFileName
  End With
  ReDim Preserve uarrDocRelacionado(UBound(uarrDocRelacionado) + 1)
End Sub

Public Function getDescription(ByVal intIndex As Integer) As String
'--> Devuelve la descripción de un documento
  getDescription = uarrDocRelacionado(intIndex).strDescription
End Function

Public Function getFileName(ByVal intIndex As Integer) As String
'--> Devuelve el nombre de fichero de un documento
  getFileName = uarrDocRelacionado(intIndex).strFileName
End Function

Public Function Count() As Integer
'--> Devuelve el número de elementos
  Count = UBound(uarrDocRelacionado)
End Function

Public Function writeHTML(ByVal uObjHTML As ClassHTML) As String
'--> Devuelve la cadena HTML de la parte del frame perteneciente a los documentos relacionados
Dim strHTML As String
Dim intIndice As Integer

  strHTML = ""
  For intIndice = 0 To UBound(uarrDocRelacionado) - 1
    With uarrDocRelacionado(intIndice)
      strHTML = strHTML + "<TR>" + uObjHTML.getHipertexto(.strFileName, .strDescription, "propiedades") + _
                          "</TR>"
    End With
  Next intIndice
  If strHTML <> "" Then strHTML = uObjHTML.getTable("Comentarios", strHTML, False)
  writeHTML = strHTML
End Function

Private Sub Class_Initialize()
  Clear
End Sub

Private Sub Class_Terminate()
  Erase uarrDocRelacionado
End Sub
