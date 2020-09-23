VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProyecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Definición de la clase proyecto
Option Explicit

Private strNombre As String 'Nombre del proyecto
Private strFicheroProyecto As String 'Fichero del proyecto
Private strDescripcion As String 'Descripción
Private strDirectorioDocumentacion As String 'Directorio donde se genera la documentación
Private blnConListaParametros As Boolean 'Si se desea o no escribir la tabla de parámetros en la documentación
Private blnConVariables As Boolean 'Si se desea o no escribir las variables locales en la documentación

Private cFormularios As New ColFormularios 'Colección con los formularios (módulos, formularios, clases...) del proyecto
Private clObjetos As New ClassObjetoReferencia 'Clase con los objetos y referencias
Private clParametros As New ClassParProyecto 'Clase con los parámetros

Private clGeneradorHTML As New ClassHTML

Public Property Let Nombre(ByVal pstrNombre As String)
  strNombre = pstrNombre
End Property

Public Property Get Nombre() As String
  Nombre = strNombre
End Property

Public Property Let Descripcion(ByVal pstrDescripcion As String)
  strDescripcion = pstrDescripcion
End Property

Public Property Get Descripcion() As String
  Descripcion = strDescripcion
End Property

Public Property Let DirectorioDocumentacion(ByVal pstrDirectorioDocumentacion As String)
  strDirectorioDocumentacion = pstrDirectorioDocumentacion
End Property

Public Property Get DirectorioDocumentacion() As String
  DirectorioDocumentacion = strDirectorioDocumentacion
End Property

Public Property Let ConParametros(ByVal blnNewConParametros As Boolean)
  blnConListaParametros = blnNewConParametros
End Property

Public Property Get ConParametros() As Boolean
  ConParametros = blnConListaParametros
End Property

Public Property Let ConVariables(ByVal blnNewConVariables As Boolean)
  blnConVariables = blnNewConVariables
End Property

Public Property Get ConVariables() As Boolean
  ConVariables = blnConVariables
End Property

Private Function getDirectorio(ByVal FicheroProyecto As String) As String
'--> Obtiene el directorio del proyecto
  On Error GoTo ErrorGetDirectorio
  getDirectorio = FicheroProyecto
  While Right$(getDirectorio, 1) <> "\" And getDirectorio <> ""
    getDirectorio = Left$(getDirectorio, Len(getDirectorio) - 1)
  Wend
  Exit Function
  
ErrorGetDirectorio:
  getDirectorio = ""
End Function

Public Function LeerProyecto(ByVal pstrNombreFichero As String) As Boolean
'--> Lee el fichero de proyecto (vbp) y va llamando a la colección de fomularios para leer cada uno de los ficheros
'--> @sub addFormulario
'--> @sub addObjetoReferencia
'--> @sub addOtros
Dim NumeroFichero As Long
Dim CadenaLectura As String
Dim Identificador As String

  On Error GoTo ErrorLectura
  cFormularios.Clear
  clObjetos.Clear
  clParametros.Clear
  strFicheroProyecto = pstrNombreFichero
  NumeroFichero = FreeFile
  Open strFicheroProyecto For Input As NumeroFichero
    Line Input #NumeroFichero, CadenaLectura
    While Not EOF(NumeroFichero)
      If Trim(CadenaLectura) <> "" Then
        Identificador = guObjGeneral.QuitarParametro(CadenaLectura, "=")
        Select Case UCase$(Identificador)
          Case "FORM", "USERCONTROL", "MODULE", "CLASS", "USERDOCUMENT"
            addFormulario Identificador, CadenaLectura
          Case "REFERENCE", "OBJECT"
            addObjetoReferencia Identificador, CadenaLectura
          Case Else
            addOtros Identificador, CadenaLectura
        End Select
      End If
      Line Input #NumeroFichero, CadenaLectura
    Wend
  Close NumeroFichero
  cFormularios.readFormularios getDirectorio(strFicheroProyecto)
  Exit Function
  
ErrorLectura:
  LeerProyecto = False
End Function

Private Sub addFormulario(ByVal Identificador As String, ByVal Cadena As String)
'--> Añade un formulario a la colección
Dim Tipo As eTipo
Dim NombreFormulario As String
Dim NombreFichero As String

  Identificador = UCase$(Identificador)
  Select Case Identificador
    Case "FORM"
      Tipo = eTipoFormulario
    Case "USERCONTROL"
      Tipo = eTipoControlUsuario
    Case "MODULE"
      Tipo = eTipoModulo
    Case "CLASS"
      Tipo = eTipoClase
    Case "USERDOCUMENT"
      Tipo = eTipoDocumentoUsuario
  End Select
  NombreFormulario = guObjGeneral.QuitarParametro(Cadena, ";")
  NombreFichero = Cadena
  If NombreFichero = "" Then NombreFichero = NombreFormulario
  cFormularios.Add NombreFormulario, NombreFichero, Tipo
End Sub

Private Sub addObjetoReferencia(ByVal Identificador As String, ByVal Cadena As String)
'--> Añade un objeto o referencia
  clObjetos.Add Identificador, Cadena
End Sub

Private Sub addOtros(ByVal Identificador As String, ByVal Cadena As String)
'--> Añade otros parámetros del proyecto
  clParametros.Add Identificador, Cadena
End Sub

Public Sub writeHTML(ByVal strFicheroIndice As String, ByVal uobjDocRelacionados As ClassDocumentado)
'--> Escribe el HTML
'--> Escribe tres ficheros:
'--> <BLOCKQUOTE>
'--> <B> Main_Doc.html </B> con la llamada a los frames
'--> <B> Main_Frame.html </B> con el frame izquierdo con todos los componentes del proyecto
'--> <B> xxxx.html </B> con el fichero índice
'--> </BLOCQUOTE>
'--> @param strFicheroIndice Fichero con la descripción del proyecto que actúa como índice (path + \ + Nombre_Fichero
'--> @param uobjColRelacionados Objeto con la colección de documentos relacionados
Dim NumeroFichero As Long
Dim ObjFichero As New ClassFicheros

  clGeneradorHTML.FicheroProyecto = strFicheroIndice
  clGeneradorHTML.writeStyleCSS strDirectorioDocumentacion
  'Genera el archivo Main_Doc.html con los frames
  NumeroFichero = FreeFile
  Open strDirectorioDocumentacion + "main_doc.html" For Output As NumeroFichero
    Print #NumeroFichero, clGeneradorHTML.getCabeceraHTML("Documentación", True)
    Print #NumeroFichero, "<FRAMESET cols=" + gstrComillas + "20%, 80%; " + gstrComillas + ">"
    Print #NumeroFichero, "<FRAME src=" + gstrComillas + "main_Frame.html" + gstrComillas + " name=" + gstrComillas + "index" + gstrComillas + "> "
    Print #NumeroFichero, "<FRAME src=" + gstrComillas + ObjFichero.ObtenerNombreFichero(strFicheroIndice) + gstrComillas + " name=" + gstrComillas + "propiedades" + gstrComillas + ">"
    Print #NumeroFichero, "</FRAMESET>"
    Print #NumeroFichero, "<NOFRAMES>"
    Print #NumeroFichero, "<H2>Error</H2>"
    Print #NumeroFichero, "<P>"
    Print #NumeroFichero, "Este documento se diseñó para verse utilizando frames, si ves este mensaje tu explorador no posee esta posibilidad."
    Print #NumeroFichero, "<BR>"
    Print #NumeroFichero, "Mirar en " + clGeneradorHTML.getHipertexto(ObjFichero.ObtenerNombreFichero(strFicheroIndice), "Versión sin frames.") + "</NOFRAMES>"
    Print #NumeroFichero, clGeneradorHTML.getCierreHTML(True)
  Close NumeroFichero
  
  'Genera el archivo Main_Frame.html con el frame izquierdo
  Open strDirectorioDocumentacion + "main_frame.html" For Output As NumeroFichero
    Print #NumeroFichero, clGeneradorHTML.getCabeceraHTML("Indice")
    If Not uobjDocRelacionados Is Nothing Then
      Print #NumeroFichero, uobjDocRelacionados.writeHTML(clGeneradorHTML)
    End If
    Print #NumeroFichero, clGeneradorHTML.getTable("Indice", _
                              "<TR>" + clGeneradorHTML.getHipertexto(ObjFichero.ObtenerNombreFichero(strFicheroIndice), "Indice del proyecto", "propiedades") + _
                              "</TR>", False)
    cFormularios.writeFrameHTML clGeneradorHTML, NumeroFichero
    Print #NumeroFichero, "</BODY> </HTML>" 'No se utiliza getCierreHTML porque es un frame
  Close NumeroFichero
  
  'Genera el fichero del proyecto Index.html
  NumeroFichero = FreeFile
  Open strFicheroIndice For Output As NumeroFichero
    Print #NumeroFichero, clGeneradorHTML.getCabeceraHTML("Proyecto " + Nombre + " - Documentación")
    'Obtiene la barra al comienzo
    Print #NumeroFichero, clGeneradorHTML.getBarraNavegadorHTML(False, True, False, False, False)
    'Obtiene los parámetros del proyecto
    Print #NumeroFichero, "<H2> " + clGeneradorHTML.getCadenaConColor("FF0000", "Proyecto: ") + _
                          strNombre + "</H2>"
    Print #NumeroFichero, "<H2> " + clGeneradorHTML.getCadenaConColor("FF0000", "Descripción: ") + _
                          strDescripcion + "</H2>"
    Print #NumeroFichero, "<H3> " + clGeneradorHTML.getCadenaConColor("FF0000", "Fichero: ") + _
                          strFicheroProyecto + "</H3>"
    Print #NumeroFichero, "<H3> " + clGeneradorHTML.getCadenaConColor("FF0000", "Fecha documentación: ") + _
                          Format(Now, "dd-mm-yyyy") + " (" + Format(Now, "hh:mm") + ") </H3>"
    Print #NumeroFichero, "<HR>"
    clObjetos.writeHTML clGeneradorHTML, NumeroFichero
    clParametros.writeHTML clGeneradorHTML, NumeroFichero
    cFormularios.writeIndexHTML clGeneradorHTML, NumeroFichero
    'Obtiene la barra al final
    Print #NumeroFichero, clGeneradorHTML.getBarraNavegadorHTML(False, True, False, False, False)
    Print #NumeroFichero, clGeneradorHTML.getCierreHTML()
  Close NumeroFichero
  'Genera los ficheros de cada uno de los formularios, clases, controles, etc...
  cFormularios.writeHTML strDirectorioDocumentacion, blnConListaParametros, blnConVariables, strNombre, clGeneradorHTML
  Set ObjFichero = Nothing
End Sub
