VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColFormularios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassFormulario"
Attribute VB_Ext_KEY = "Member0" ,"ClassFormulario"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección de formularios
Option Explicit

Private ColeccionFormularios As Collection 'Colección con objetos de tipo ClassFormulario

Public Sub Clear()
'--> Elimina los formularios de la colección
Dim intIndice As Integer

  For intIndice = ColeccionFormularios.Count To 1 Step -1
    ColeccionFormularios.Remove intIndice
  Next intIndice
End Sub

Public Function Add(ByVal NombreFormulario As String, ByVal NombreFichero As String, _
                    ByVal Tipo As eTipo, Optional ByVal strKey As String) As ClassFormulario
'--> Crea un nuevo formulario
Dim objNewFormulario As ClassFormulario

  Set objNewFormulario = New ClassFormulario
  objNewFormulario.NombreFormulario = NombreFormulario
  objNewFormulario.NombreFichero = NombreFichero
  objNewFormulario.Tipo = Tipo
  If Len(strKey) = 0 Then
      ColeccionFormularios.Add objNewFormulario
  Else
      ColeccionFormularios.Add objNewFormulario, strKey
  End If
  Set Add = objNewFormulario
  Set objNewFormulario = Nothing
End Function

Public Property Get Item(Index As Variant) As ClassFormulario
Attribute Item.VB_UserMemId = 0
'--> Obtiene el formulario index de la colección
  Set Item = ColeccionFormularios(Index)
End Property

Public Property Get Count() As Long
'--> Obtiene el número de elementos en la colección (en este caso el número de formularios)
  Count = ColeccionFormularios.Count
End Property

Public Sub Remove(Index As Variant)
'--> Borra un elemento de la colección
  ColeccionFormularios.Remove Index
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'--> Propiedad necesaria en VB para las estructuras del tipo for each ...
  Set NewEnum = ColeccionFormularios.[_NewEnum]
End Property

Public Sub readFormularios(ByVal Directorio As String)
'--> Lee los formularios (llama al método Read de la clase formulario)
Dim clFormulario As ClassFormulario

  For Each clFormulario In ColeccionFormularios
    clFormulario.read Directorio
  Next clFormulario
End Sub

Public Function getIndex(ByRef clGeneradorHTML As ClassHTML, ByVal Tipo As eTipo) As String
'--> Obtiene el index HTML de un tipo de formulario
Dim clFormulario As ClassFormulario
Dim CadSalida As String
Dim Indice As Integer

  Indice = 0
  CadSalida = "<TR>"
  CadSalida = CadSalida + "<TD CLASS=TableSubcaption> Nombre"
  CadSalida = CadSalida + "<TD CLASS=TableSubcaption> Fichero"
  For Each clFormulario In ColeccionFormularios
    If clFormulario.Tipo = Tipo Then
      CadSalida = CadSalida + "<TR>"
      CadSalida = CadSalida + "<TD> " + clGeneradorHTML.getHipertexto(clFormulario.NombreFormulario + ".html", _
                                                                      clFormulario.NombreFormulario)
      CadSalida = CadSalida + "<TD>" + clFormulario.NombreFichero
      Indice = Indice + 1
    End If
  Next clFormulario
  getIndex = IIf(Indice = 0, "", CadSalida) '+ "</TABLE>")
End Function

Public Sub writeIndexHTML(ByRef clGeneradorHTML As ClassHTML, ByVal NumeroFichero As Long)
'--> Escribe el HTML con los nombres de los formularios
Dim Indice As Integer
Dim CadSalida As String, CadFormulario As String

  CadSalida = ""
  CadFormulario = getIndex(clGeneradorHTML, eTipoModulo)
  If CadFormulario <> "" Then
    CadSalida = CadSalida + clGeneradorHTML.getTable("Módulos", CadFormulario)
  End If
  CadFormulario = getIndex(clGeneradorHTML, eTipoFormulario)
  If CadFormulario <> "" Then
    CadSalida = CadSalida + clGeneradorHTML.getTable("Formularios", CadFormulario)
  End If
  CadFormulario = getIndex(clGeneradorHTML, eTipoClase)
  If CadFormulario <> "" Then
    CadSalida = CadSalida + clGeneradorHTML.getTable("Clases", CadFormulario)
  End If
  CadFormulario = getIndex(clGeneradorHTML, eTipoControlUsuario)
  If CadFormulario <> "" Then
    CadSalida = CadSalida + clGeneradorHTML.getTable("Controles de Usuario", CadFormulario)
  End If
  CadFormulario = getIndex(clGeneradorHTML, eTipoDocumentoUsuario)
  If CadFormulario <> "" Then
    CadSalida = CadSalida + clGeneradorHTML.getTable("Documentos de usuario", CadFormulario)
  End If
  Print #NumeroFichero, CadSalida + "<HR>"
End Sub

Public Function getIndexFrame(ByRef clGeneradorHTML As ClassHTML, ByVal Tipo As eTipo) As String
'--> Obtiene el index del frame de un tipo de formulario
Dim clFormulario As ClassFormulario
Dim CadSalida As String

  CadSalida = ""
  For Each clFormulario In ColeccionFormularios
    If clFormulario.Tipo = Tipo Then
      CadSalida = CadSalida + "<TR>" + clGeneradorHTML.getHipertexto(clFormulario.NombreFormulario + ".html", _
                                                                     clFormulario.NombreFormulario, "propiedades") + _
                              " </TR>"
    End If
  Next clFormulario
  getIndexFrame = CadSalida
End Function

Public Sub writeFrameHTML(ByRef clGeneradorHTML As ClassHTML, ByVal NumeroFichero As Long)
'--> Escribe el frame izquierdo con los nombres de los formularios
'--> @sub getIndexFrmae
Dim Indice As Integer
Dim CadSalida As String, CadFormulario As String

  CadSalida = ""
  CadFormulario = getIndexFrame(clGeneradorHTML, eTipoModulo)
  If CadFormulario <> "" Then
    CadSalida = clGeneradorHTML.getTable("Módulos", CadFormulario, False)
  End If
  CadFormulario = getIndexFrame(clGeneradorHTML, eTipoFormulario)
  If CadFormulario <> "" Then
    CadSalida = CadSalida + clGeneradorHTML.getTable("Formularios", CadFormulario, False)
  End If
  CadFormulario = getIndexFrame(clGeneradorHTML, eTipoClase)
  If CadFormulario <> "" Then
    CadSalida = CadSalida + clGeneradorHTML.getTable("Clases", CadFormulario, False)
  End If
  CadFormulario = getIndexFrame(clGeneradorHTML, eTipoControlUsuario)
  If CadFormulario <> "" Then
    CadSalida = CadSalida + clGeneradorHTML.getTable("Controles de Usuario", CadFormulario, False)
  End If
  CadFormulario = getIndexFrame(clGeneradorHTML, eTipoDocumentoUsuario)
  If CadFormulario <> "" Then
    CadSalida = CadSalida + clGeneradorHTML.getTable("Documentos de Usuario", CadFormulario, False)
  End If
  Print #NumeroFichero, CadSalida
End Sub

Public Sub writeHTML(ByVal Directorio As String, ByVal blnConListaParametros As Boolean, ByVal blnConVariables As Boolean, _
                     ByVal strProyecto As String, ByRef clGeneradorHTML As ClassHTML)
'--> Escribe los HTML de cada formulario, para ello recorre la colección de formularios llamando al método writeHTML de la clase formulario
Dim clFormulario As ClassFormulario

  For Each clFormulario In ColeccionFormularios
    clFormulario.writeHTML Directorio, blnConListaParametros, blnConVariables, strProyecto, clGeneradorHTML
  Next clFormulario
End Sub

Private Sub Class_Initialize()
  Set ColeccionFormularios = New Collection
End Sub

Private Sub Class_Terminate()
  Set ColeccionFormularios = Nothing
End Sub
