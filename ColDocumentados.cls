VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColDocumentados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassDocumentado"
Attribute VB_Ext_KEY = "Member0" ,"ClassDocumentado"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección con los proyectos ya documentados
Option Explicit

Private cColDocumentos As Collection 'Colección para los documentos de tipo ClassDocumentado

Public Sub readIni(ByVal strAppPath As String)
'--> Lee el fichero ini que está en App.Path + "\Documentador.ini"
'--> @param strAppPath Path de la aplicación
Dim strNombreFichero As String
Dim intIndice As Integer, intNumProyectos As Integer
Dim strDescription As String, strName As String, strDateGeneration As String
Dim strFileIndex As String, strFileProject As String
Dim blnWithVariables As Boolean, blnWithParameters As Boolean
Dim intNumeroRelacionados As Integer, intRelacionado As Integer
Dim strRelacionado As String
Dim objDocumentado As ClassDocumentado

  strNombreFichero = strAppPath + "\Documentador.ini"
  intNumProyectos = Val(guObjGeneral.CargarCadenaIni(strNombreFichero, "Documentador", "NumeroProyectos", "0"))
  For intIndice = 1 To intNumProyectos
    strName = guObjGeneral.CargarCadenaIni(strNombreFichero, "Fichero" & intIndice, "Nombre", "0")
    strDescription = guObjGeneral.CargarCadenaIni(strNombreFichero, "Fichero" & intIndice, "Descripcion", "")
    strFileProject = guObjGeneral.CargarCadenaIni(strNombreFichero, "Fichero" & intIndice, "DirProyecto", "")
    strFileIndex = guObjGeneral.CargarCadenaIni(strNombreFichero, "Fichero" & intIndice, "IndexDocumentacion", "")
    blnWithParameters = (UCase(guObjGeneral.CargarCadenaIni(strNombreFichero, "Fichero" & intIndice, "ConParametros", "No")) <> "NO")
    blnWithVariables = (UCase(guObjGeneral.CargarCadenaIni(strNombreFichero, "Fichero" & intIndice, "ConVariable", "No")) <> "NO")
    strDateGeneration = guObjGeneral.CargarCadenaIni(strNombreFichero, "Fichero" & intIndice, "FechaGeneracion", Format(Now, ConstCadenaFecha))
    Set objDocumentado = Add(blnWithVariables, blnWithParameters, strDateGeneration, _
                             strFileIndex, strFileProject, strDescription, strName, Nothing)
    objDocumentado.Clear
    intNumeroRelacionados = Val(guObjGeneral.CargarCadenaIni(strNombreFichero, "Fichero" & intIndice, "NumRelacionados", "0"))
    For intRelacionado = 1 To intNumeroRelacionados
      strRelacionado = Trim(guObjGeneral.CargarCadenaIni(strNombreFichero, "Fichero" & intIndice, "DocRel" & intRelacionado, ""))
      If strRelacionado <> "" Then
        objDocumentado.addItem guObjGeneral.QuitarParametro(strRelacionado, "@"), strRelacionado
      End If
    Next intRelacionado
    Set objDocumentado = Nothing
  Next intIndice
End Sub

Public Sub writeIni(ByVal strAppPath As String)
'--> Escribe el fichero ini que está en App.Path + "\Documentador.ini"
'--> @param strAppPath Path de la aplicación
Dim NumeroFichero As Long
Dim objDocumentado As ClassDocumentado
Dim intIndice As Integer, intRelacionado As Integer

  NumeroFichero = FreeFile
  Open strAppPath + "\Documentador.ini" For Output As #NumeroFichero
    Print #NumeroFichero, "[Documentador]"
    Print #NumeroFichero, "NumeroProyectos=" + Str(cColDocumentos.Count)
    Print #NumeroFichero, ""
    
    intIndice = 1
    For Each objDocumentado In cColDocumentos
      Print #NumeroFichero, "[Fichero" + Trim(Str(intIndice)) + "]"
      With objDocumentado
        Print #NumeroFichero, "Nombre=" + .strName
        Print #NumeroFichero, "Descripcion=" + .strDescription
        Print #NumeroFichero, "DirProyecto=" + .strFileProject
        Print #NumeroFichero, "IndexDocumentacion=" + .strFileIndex
        Print #NumeroFichero, "ConParametros=" + IIf(.blnWithParameters, "SI", "NO")
        Print #NumeroFichero, "ConVariable=" + IIf(.blnWithVariables, "SI", "NO")
        Print #NumeroFichero, "FechaGeneracion=" + .strDateGeneration
        Print #NumeroFichero, "NumRelacionados=" + Str(objDocumentado.Count)
        For intRelacionado = 0 To objDocumentado.Count - 1
          Print #NumeroFichero, "DocRel" & (intRelacionado + 1) & "=" & objDocumentado.getDescription(intRelacionado) & "@" & _
                                                                        objDocumentado.getFileName(intRelacionado)
        Next intRelacionado
      End With
      Print #NumeroFichero, ""
      intIndice = intIndice + 1
    Next objDocumentado
  Close #NumeroFichero
End Sub

Public Sub loadList(ByVal lstProyectos As ListBox)
'--> Carga una lista con todos los ficheros documentados
'--> @param lstProyectos Lista donde debe almacenar los proyectos
Dim objDocumentado As ClassDocumentado

  lstProyectos.Clear
  For Each objDocumentado In cColDocumentos
    lstProyectos.addItem objDocumentado.strName
  Next objDocumentado
End Sub

Public Sub loadGrid(ByVal intItem As Integer, ByVal grdDatos As MSFlexGrid)
'--> Carga el grid con los documentos relacionados al proyecto
Dim objDocumentado As ClassDocumentado
Dim intIndice As Integer

  On Error GoTo ErrorCarga
  Set objDocumentado = cColDocumentos.Item(intItem)
  With grdDatos
    'Borra los datos anteriores
    .Rows = 2
    .TextMatrix(1, 0) = ""
    .TextMatrix(1, 1) = ""
    'Añade los elementos relacionados
    For intIndice = 0 To objDocumentado.Count - 1
      .addItem objDocumentado.getDescription(intIndice) & vbTab & objDocumentado.getFileName(intIndice)
    Next intIndice
    If .Rows > 2 Then .RemoveItem 1
  End With
  Set objDocumentado = Nothing
  Exit Sub
  
ErrorCarga:
End Sub

Private Sub addDocRelacionados(ByRef objDocumentado As ClassDocumentado, ByRef grdDocumentos As MSFlexGrid)
'--> Añade los documentos relacionados del grid al objeto
Dim intIndice As Integer

  If Not grdDocumentos Is Nothing Then
    objDocumentado.Clear
    With grdDocumentos
      For intIndice = 1 To grdDocumentos.Rows - 1
        If Trim(.TextMatrix(intIndice, 0)) <> "" Then
          objDocumentado.addItem .TextMatrix(intIndice, 0), .TextMatrix(intIndice, 1)
        End If
      Next intIndice
    End With
  End If
End Sub

Public Sub Modify(ByVal intIndex As Integer, ByVal blnWithVariables As Boolean, ByVal blnWithParameters As Boolean, _
                  ByVal strDateGeneration As String, ByVal strFileIndex As String, ByVal strFileProject As String, _
                  ByVal strDescription As String, ByVal strName As String, ByRef grdDocumentos As MSFlexGrid)
'--> Modifica un elemento de la coleccion, si hay algún error lo añade
Dim objDocumentado As ClassDocumentado

  On Error GoTo ErrorModificacion
  Set objDocumentado = cColDocumentos.Item(intIndex + 1)
  With objDocumentado
    .strName = strName
    .strDescription = strDescription
    .strFileProject = strFileProject
    .strFileIndex = strFileIndex
    .blnWithParameters = blnWithParameters
    .blnWithVariables = blnWithVariables
    .strDateGeneration = strDateGeneration
  End With
  addDocRelacionados objDocumentado, grdDocumentos
  Set objDocumentado = Nothing
  Exit Sub
  
ErrorModificacion:
  Add blnWithVariables, blnWithParameters, strDateGeneration, strFileIndex, strFileProject, _
      strDescription, strName, grdDocumentos
End Sub

Public Function Add(ByVal blnWithVariables As Boolean, ByVal blnWithParameters As Boolean, _
                    ByVal strDateGeneration As String, ByVal strFileIndex As String, ByVal strFileProject As String, _
                    ByVal strDescription As String, ByVal strName As String, ByRef grdDocumentos As MSFlexGrid, Optional sKey As String) As ClassDocumentado
'--> Crear un nuevo objeto
Dim objNewMember As ClassDocumentado

  Set objNewMember = New ClassDocumentado
  'establece las propiedades que se transfieren al método
  objNewMember.blnWithVariables = blnWithVariables
  objNewMember.blnWithParameters = blnWithParameters
  objNewMember.strDateGeneration = strDateGeneration
  objNewMember.strFileIndex = strFileIndex
  objNewMember.strFileProject = strFileProject
  objNewMember.strDescription = strDescription
  objNewMember.strName = strName
  addDocRelacionados objNewMember, grdDocumentos
  If Len(sKey) = 0 Then
    cColDocumentos.Add objNewMember
  Else
    cColDocumentos.Add objNewMember, sKey
  End If
  'devolver el objeto creado
  Set Add = objNewMember
  Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As ClassDocumentado
Attribute Item.VB_UserMemId = 0
'--> Obtiene un objeto de la clase
  On Error GoTo ErrorItem
  Set Item = cColDocumentos(vntIndexKey)
  Exit Property
  
ErrorItem:
  Set Item = Nothing
End Property

Public Property Get Count() As Long
'--> Obtiene el número de elementos de la colección
  Count = cColDocumentos.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
'--> Elimina un elemento
  cColDocumentos.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'--> Para las acciones For...Each
  Set NewEnum = cColDocumentos.[_NewEnum]
End Property

Private Sub Class_Initialize()
'--> Al inicializar la clase crea la colección
  Set cColDocumentos = New Collection
End Sub

Private Sub Class_Terminate()
'--> Destruye la colección cuando se termina la clase
  Set cColDocumentos = Nothing
End Sub
