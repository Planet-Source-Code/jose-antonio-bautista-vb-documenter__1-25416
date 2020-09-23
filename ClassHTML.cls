VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase interna de ayuda para generar ficheros HTML
Option Explicit

Private strFicheroProyecto As String

Public Property Let FicheroProyecto(ByVal strNewFicheroProyecto As String)
  strFicheroProyecto = strNewFicheroProyecto
End Property

Public Property Get FicheroProyecto() As String
  FicheroProyecto = strFicheroProyecto
End Property

Public Sub writeStyleCSS(ByVal strDirectorioDocumentacion As String)
'--> Escribe el fichero de estilos
Dim lngFichero As Long

  lngFichero = FreeFile
  Open strDirectorioDocumentacion + "Style.css" For Output As lngFichero
    Print #lngFichero, "/* Documentador HTML style sheet */"
    Print #lngFichero, "/* Define colores, fuentes y atributos de estilo  */"
    
    Print #lngFichero, "/* Background de página */"
    Print #lngFichero, "body { background-color: #FFFFFF }"
    
    Print #lngFichero, "/* Colores de los enlaces */"
    Print #lngFichero, "A:active { COLOR: navy; Text-Decoration: Underline }"
    Print #lngFichero, "A:link { COLOR: navy }"
    Print #lngFichero, "A:visited { COLOR: olive }"
    Print #lngFichero, "A {COLOR: navy }"
    
    Print #lngFichero, "/* Colores generales */"
    Print #lngFichero, ".Header { background: #EEEEEE; font-size: 17; FONT-WEIGHT: bolder; color: #0000FF;}"
    Print #lngFichero, ".SubroutineHeader { COLOR: #3333cc; FONT-SIZE: 15; FONT-WEIGHT: bolder; Text-Decoration: Underline;}"
       
    Print #lngFichero, "/* Colores de tabla */"
    Print #lngFichero, ".TableCaption { background: #EEEEEE; font-size: 15; color: #0000FF;}"
    Print #lngFichero, ".TableSubcaption { background: #EEEEFF; font-size: 12; FONT-WEIGHT: bolder; color: #AA00AA;}"
    Print #lngFichero, ".TableRowColor { background: #FFFFFF; } /* White */"
    Print #lngFichero, ".RowHeader { background: #EEEEFF;}"
    
    Print #lngFichero, "/* Barra de navegación */"
    Print #lngFichero, ".BotonEnabled {background-color:#EEEEFF; color:#0000FF;}"
    Print #lngFichero, ".BotonDisabled {background-color:#CCCCFF; color:#000000;}"
    
  Close lngFichero
End Sub

Public Function getCabeceraHTML(ByVal Titulo As String, Optional ByVal blnWithFrames As Boolean = False) As String
'--> Devuelve una cabecera válida HTML
  getCabeceraHTML = "<HTML>"
  getCabeceraHTML = getCabeceraHTML + "<HEAD>"
  getCabeceraHTML = getCabeceraHTML + "<META NAME=""""; GENERATOR; """"; CONTENT = """"" + App.ProductName + """"" > "
  getCabeceraHTML = getCabeceraHTML + "<META HTTP-EQUIV=""""Content-Type""""; CONTENT = """"; Text / html; Charset = iso - 8859 - 1; """" > "
  getCabeceraHTML = getCabeceraHTML + "<!-- Generado por Documentador HTML - " + Format(Now, "dd-mm-yyyy") + "-->"
  getCabeceraHTML = getCabeceraHTML + "<TITLE> " + Titulo + "</TITLE>"
  getCabeceraHTML = getCabeceraHTML + "<LINK REL =" + gstrComillas + "stylesheet" + gstrComillas + _
                                      " TYPE=" + gstrComillas + "text/css" + gstrComillas + _
                                      " HREF=" + gstrComillas + "Style.css" + gstrComillas + _
                                      " TITLE=" + gstrComillas + "Style" + gstrComillas + ">"
  getCabeceraHTML = getCabeceraHTML + "</HEAD>"
  If Not blnWithFrames Then
    getCabeceraHTML = getCabeceraHTML + "<BODY>"
  End If
End Function

Private Function getBotonNavegador(ByVal blnEnabled As Boolean, ByVal strURL As String, ByVal strTitle As String) As String
'--> Obtiene un botón de la barra de menús del documento HTML activo o inactivo
  getBotonNavegador = "<TD CLASS=" + IIf(blnEnabled, "BotonEnabled", "BotonDisabled") + _
                      "> " + IIf(blnEnabled, getHipertexto(strURL, strTitle), strTitle) + _
                      " </TD>"
End Function

Public Function getBarraNavegadorHTML(ByVal blnWithIndex As Boolean, ByVal blnDescripcion As Boolean, _
                                      ByVal blnControles As Boolean, ByVal blnVariables As Boolean, _
                                      ByVal blnRutinas As Boolean) As String
'--> Obtiene la barra del navegador completa
'--> @sub getBotonNavegador
  getBarraNavegadorHTML = getBarraNavegadorHTML + "<TABLE BORDER=0 WIDTH=100% CELLPADDING=1 CELLSPACING=0>"
  getBarraNavegadorHTML = getBarraNavegadorHTML + "<TR>"
  getBarraNavegadorHTML = getBarraNavegadorHTML + getBotonNavegador(blnWithIndex, FicheroProyecto, "Indice")
  getBarraNavegadorHTML = getBarraNavegadorHTML + getBotonNavegador(blnDescripcion, "#Descripcion", "Descripción")
  getBarraNavegadorHTML = getBarraNavegadorHTML + getBotonNavegador(blnControles, "#Controles", "Controles")
  getBarraNavegadorHTML = getBarraNavegadorHTML + getBotonNavegador(blnVariables, "#Variables", "Variables")
  getBarraNavegadorHTML = getBarraNavegadorHTML + getBotonNavegador(blnRutinas, "#Rutinas", "Rutinas")
  getBarraNavegadorHTML = getBarraNavegadorHTML + "</TR>"
  getBarraNavegadorHTML = getBarraNavegadorHTML + "</TABLE>"
  getBarraNavegadorHTML = getBarraNavegadorHTML + "<HR><BR>"
End Function

Public Function getCierreHTML(Optional ByVal blnWithFrames As Boolean = False) As String
'--> Obtiene el cierre de un fichero HTML
  getCierreHTML = ""
  If Not blnWithFrames Then
    getCierreHTML = "<font size=-1>"
    getCierreHTML = getCierreHTML + "Para cualquier error o sugerencia diríjase a :" + _
                    getHipertexto("http://hhhh", "Gerencia de Informática - CEI")
    getCierreHTML = getCierreHTML + "<BR> @1998 - 2000 </font>"
    getCierreHTML = getCierreHTML + "</BODY>"
  End If
  getCierreHTML = getCierreHTML + "</HTML>"
End Function

Public Function getCadenaConColor(ByVal COLOR As String, ByVal Cadena As String) As String
'--> Obtiene una cadena HTML con color
  getCadenaConColor = "<FONT COLOR=" + COLOR + "> " + Cadena + "</FONT>"
End Function

Public Function getCadenaNegrita(ByVal Cadena As String) As String
'--> Obtiene una cadena HTML en negrita
  getCadenaNegrita = "<STRONG> " + Cadena + " </STRONG>"
End Function

Public Function getHipertexto(ByVal strURL As String, ByVal strCaption As String, Optional ByVal strFrame As String = "") As String
'--> Obtiene una cadena de URL
  If strFrame = "" Then
    getHipertexto = "<A HREF=" + gstrComillas + strURL + gstrComillas + ">" + strCaption + "</A>"
  Else
    getHipertexto = "<A HREF=" + gstrComillas + strURL + gstrComillas + _
                    " TARGET=" + gstrComillas + strFrame + gstrComillas + ">" + strCaption + "</A>"
  End If
End Function

Public Function getAnchor(ByVal strCaption As String) As String
'--> Obtiene una cadena de Anchor
  getAnchor = "<A NAME=" + gstrComillas + strCaption + gstrComillas + "></A>"
End Function

Public Function getTable(ByVal strCaption As String, ByVal strDescriptionRows As String, Optional ByVal blnConSalto As Boolean = True) As String
'--> Obtiene una tabla con el título indicado y las filas pasadas como descripción
  getTable = "<TABLE BORDER=1 CELLPADDING=3 CELLSPACING=0 WIDTH=100%>"
  getTable = getTable + "<TR ALIGN=Center BORDER=1 CLASS=TableCaption>" + strCaption + " </TR>"
  getTable = getTable + strDescriptionRows
  getTable = getTable + "</TABLE>"
  If blnConSalto Then
    getTable = getTable + "<HR> <BR>"
  Else
    getTable = getTable + "&nbsp"
  End If
End Function

Public Function getCadenaVacia(ByVal strCaption As String) As String
'--> Si strCaption viene vacía se pone un espacio fijo, útil en las tablas cuando deseamos que nos ponga los bordes _
     aunque la celda no tenga contenido
  getCadenaVacia = IIf(Trim(strCaption) = "", "&nbsp", strCaption)
End Function
