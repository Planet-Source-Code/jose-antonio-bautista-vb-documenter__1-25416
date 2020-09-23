VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFicheros"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase con rutinas útiles de control de ficheros
Option Explicit

Public Function ExisteFichero(ByVal NombreFichero As String) As Boolean
'--> Comprueba si existe el fichero
Dim LongFichero As Long

  On Error GoTo ErrorExisteFichero
  ExisteFichero = True
  LongFichero = FileLen(NombreFichero)
  Exit Function
  
ErrorExisteFichero:
  ExisteFichero = False
End Function

Public Function ObtenerPath(ByVal NombreFichero As String) As String
'--> A partir de un nombre de fichero consigue el directorio a quien pertenece
'--> @param NombreFichero Nombre completo del fichero (pe: C:\TEMP\Documento.doc
'--> @return Path (pe: c:\TEMP)
Dim Directorio As String

  Directorio = NombreFichero
  While Right$(Directorio, 1) <> "\" And Len(Directorio) > 1
    Directorio = Left$(Directorio, Len(Directorio) - 1)
  Wend
  ObtenerPath = Directorio
End Function

Public Function ObtenerNombreFichero(ByVal NombreFichero As String) As String
'--> Obtiene el nombre del fichero del path total introducido
'--> @param NombreFichero Nombre completo del fichero (pe: C:\TEMP\Documento.doc
'--> @return Nombre del fichero (pe: Documento.doc)
Dim Indice As Integer
Dim Encontrado As Boolean

  Indice = Len(NombreFichero)
  Encontrado = False
  While Not Encontrado And Indice > 0
    If Mid$(NombreFichero, Indice, 1) = "\" Then Encontrado = True
    Indice = Indice - 1
  Wend
  If Encontrado Then
    ObtenerNombreFichero = Mid$(NombreFichero, Indice + 2, Len(NombreFichero))
  Else
    ObtenerNombreFichero = NombreFichero
  End If
End Function

Public Sub BorrarFichero(ByVal NombreFichero As String)
'--> Borra un fichero
  On Error Resume Next
  Kill NombreFichero
End Sub

Public Function DLGNombreFichero(DLG_Ficheros As Object, _
                                 Optional ByVal CargarGrabar As Boolean = True, _
                                 Optional ByVal NombreFichero As String = "C:\", _
                                 Optional ByVal Filtro As String = "Todos los ficheros (*.*) | *.*") As String
'--> Abre un diálogo para conseguir un nombre de fichero
  On Error GoTo ErrorApertura
  With DLG_Ficheros
    .CancelError = True
    If NombreFichero <> "" Then
      .InitDir = ObtenerPath(NombreFichero)
      .FileName = ObtenerNombreFichero(NombreFichero)
    Else
      .InitDir = "C:\"
      .FileName = ""
    End If
    .Filter = Filtro
    .FilterIndex = 0
    If CargarGrabar Then
      .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNLongNames Or _
                           cdlOFNAllowMultiselect Or cdlOFNPathMustExist Or cdlOFNNoReadOnlyReturn
      .DialogTitle = "Abrir Archivos"
      .ShowOpen
    Else
      .Flags = cdlOFNExplorer Or cdlOFNLongNames Or _
                           cdlOFNAllowMultiselect Or cdlOFNPathMustExist
      .DialogTitle = "Guardar Archivo"
      .ShowSave
    End If
    DLGNombreFichero = DLG_Ficheros.FileName
  End With
  Exit Function
  
ErrorApertura:
  DLGNombreFichero = ""
End Function
