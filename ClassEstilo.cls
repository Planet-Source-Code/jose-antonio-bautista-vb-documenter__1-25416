VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEstilo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> Definición de los tags de estilo
'--> Se utilizaba para dejar al usuario cambiar el fichero CSS, ya no funciona, hay un CSS fijo
Option Explicit

Public Enum eTagNombre
  eTagBody = 0
  eTagA
  eTagALink
  eTagAActive
  eTagAVisited
  eTagHeader
  eTagSubroutineHeader
  eTagTableCaption
  eTagTableSubCaption
  eTagTableRowColor
  eTagRowHeader
  eTagBotonEnabled
  eTagBotonDisabled
End Enum

Private Type tEstilo
  uTagNombre As eTagNombre
  strName As String
  colFondo As OLE_COLOR
  colTexto As OLE_COLOR
  blnBold As Boolean
  blnUnderline As Boolean
  intSize As Integer
End Type

Private arrTagEstilos() As tEstilo

Public Sub loadComboNombres(ByVal cboTags As ComboBox)
'--> Carga un combo con los tags definidos
Dim intIndice As Integer
  
  cboTags.Clear
  For intIndice = LBound(arrTagEstilos) To UBound(arrTagEstilos)
    cboTags.addItem arrTagEstilos(intIndice).strName
  Next intIndice
  If cboTags.ListCount > 0 Then cboTags.ListIndex = 0
End Sub

Public Sub loadComboSize(ByVal cboSize As ComboBox)
'--> Carga un combo con los tamaños
  cboSize.Clear
  cboSize.addItem "Rel. Mayor"
  cboSize.addItem "Rel. Menor"
End Sub

Private Function getColor(ByVal strColor As String) As OLE_COLOR
'--> Obtiene un color de una cadena
Dim lngRed As Long, lngGreen As Long, lngBlue As Long

  lngRed = Val(guObjGeneral.QuitarParametro(strColor, ","))
  lngGreen = Val(guObjGeneral.QuitarParametro(strColor, ","))
  lngBlue = Val(strColor)
  getColor = lngRed * 256 * 256 + lngGreen * 256 + lngBlue
End Function

Public Sub readIni(ByVal strDirectorio As String)
'--> Lee el fichero de inicializacion DocStyle.ini
Dim strFileIni As String 'Nombre del fichero .ini
Dim intIndice As Integer

  strFileIni = strDirectorio + "\DocStyle.ini"
  For intIndice = LBound(arrTagEstilos) To UBound(arrTagEstilos)
    With arrTagEstilos(intIndice)
      .colFondo = getColor(guObjGeneral.CargarCadenaIni(strFileIni, .strName, "Background", "255,255,255"))
      .colTexto = getColor(guObjGeneral.CargarCadenaIni(strFileIni, .strName, "Text", "0,0,0"))
      .blnBold = (UCase(guObjGeneral.CargarCadenaIni(strFileIni, .strName, "Bold", "No")) = "NO")
      .blnUnderline = (UCase(guObjGeneral.CargarCadenaIni(strFileIni, .strName, "Underline", "No")) = "NO")
      .intSize = Val(guObjGeneral.CargarCadenaIni(strFileIni, .strName, "Size", -1))
    End With
  Next intIndice
End Sub

Private Sub getComponentesColor(ByVal colColor As OLE_COLOR, ByRef lngRed As Long, _
                                ByRef lngGreen As Long, ByRef lngBlue As Long)
'--> Obtiene los componentes de color desde un OLE_COLOR
  lngBlue = colColor Mod 256
  colColor = Int(colColor / 256)
  lngGreen = colColor Mod 256
  colColor = Int(colColor / 256)
  lngRed = colColor
End Sub

Private Function getColorIni(ByVal colColor As OLE_COLOR) As String
'--> Obtiene una cadena de color del estilo ROJO , VERDE , AZUL para escribir sobre el INI
Dim lngRed As Long, lngGreen As Long, lngBlue As Long

  getComponentesColor colColor, lngRed, lngGreen, lngBlue
  getColorIni = Str(lngRed) + "," + Str(lngGreen) + "," + Str(lngBlue)
End Function

Public Sub writeIni(ByVal strDirectorio As String)
'--> Escribe el fichero de inicializacion DocStyle.ini
Dim strFileIni As String 'Nombre del fichero .ini
Dim lngNumFile As Long
Dim intIndice As Integer

  lngNumFile = FreeFile
  Open strDirectorio + "\DocStyle.ini" For Output As lngNumFile
  For intIndice = LBound(arrTagEstilos) To UBound(arrTagEstilos)
    With arrTagEstilos(intIndice)
      Print #lngNumFile, "[" + .strName + "]"
      Print #lngNumFile, "Background=" + getColorIni(.colFondo)
      Print #lngNumFile, "Text=" + getColorIni(.colTexto)
      Print #lngNumFile, "Bold=" + IIf(.blnBold, "Sí", "No")
      Print #lngNumFile, "Underline=" + IIf(.blnUnderline, "Sí", "No")
      Print #lngNumFile, "Size=" + Str(.intSize)
      Print #lngNumFile, ""
    End With
  Next intIndice
  Close #lngNumFile
End Sub

Public Sub changeTag(ByVal intIndiceTag As Integer, ByVal colBackGround As OLE_COLOR, ByVal colText As OLE_COLOR, _
                     ByVal blnBold As Boolean, ByVal blnUnderline As Boolean, ByVal strSize As String)
'--> Cambia un TAG de la tabla de estilos
  With arrTagEstilos(intIndiceTag)
    .colFondo = colBackGround
    .colTexto = colText
    .blnBold = blnBold
    .blnUnderline = blnUnderline
    .intSize = IIf(Trim(strSize) = "", -1, Val(strSize))
  End With
End Sub

Public Sub getTag(ByVal intIndiceTag As Integer, ByRef colBackGround As OLE_COLOR, ByRef colText As OLE_COLOR, _
                  ByRef blnBold As Boolean, ByRef blnUnderline As Boolean, ByRef strSize As String)
'--> Cambia un TAG de la tabla de estilos
  With arrTagEstilos(intIndiceTag)
    colBackGround = .colFondo
    colText = .colTexto
    blnBold = .blnBold
    blnUnderline = .blnUnderline
    strSize = IIf(.intSize = -1, " ", Trim(Str(.intSize)))
  End With
End Sub

Private Function getCharHexadecimal(ByVal intValor As Integer) As String
'--> Obtiene un valor hexadecimal 0..F
  If intValor > 9 Then
    Select Case intValor
      Case 10
        getCharHexadecimal = "A"
      Case 11
        getCharHexadecimal = "B"
      Case 12
        getCharHexadecimal = "C"
      Case 13
        getCharHexadecimal = "D"
      Case 14
        getCharHexadecimal = "E"
      Case 15
        getCharHexadecimal = "F"
    End Select
  Else
    getCharHexadecimal = Trim(Str(intValor))
  End If
End Function

Private Function getColorStyle(ByVal colColor As Long) As String
'--> Obtiene el texto de un color a grabar sobre el fichero de estilos
Dim lngRed As Long, lngGreen As Long, lngBlue As Long

  getComponentesColor colColor, lngRed, lngGreen, lngBlue
  getColorStyle = "#" + getCharHexadecimal(Int(lngRed / 16)) + getCharHexadecimal(lngRed Mod 16) + _
                        getCharHexadecimal(Int(lngGreen / 16)) + getCharHexadecimal(lngGreen Mod 16) + _
                        getCharHexadecimal(Int(lngBlue / 16)) + getCharHexadecimal(lngBlue Mod 16)
End Function

Private Function getTextStyle(ByVal intIndiceTag As eTagNombre) As String
'--> Obtiene el texto de un estilo determinado
Dim strSalida As String

  strSalida = ""
  With arrTagEstilos(intIndiceTag)
    If .colFondo <> vbWhite Then
      strSalida = "BACKGROUND-COLOR: " + getColorStyle(.colFondo) + "; "
    End If
    If .colTexto <> vbBlack Then
      strSalida = strSalida + "COLOR: " + getColorStyle(.colTexto) + "; "
    End If
    If .intSize <> -1 Then
      strSalida = strSalida + "FONT-SIZE: " + Str(.intSize) + "; "
    End If
    If .blnBold Then
      strSalida = strSalida + "FONT-WEIGHT: Bolder; "
    End If
    If .blnUnderline Then
      strSalida = strSalida + "TEXT-DECORATION: Underline; "
    End If
  End With
  If strSalida = "" Then
    getTextStyle = ""
  Else
    getTextStyle = Trim(arrTagEstilos(intIndiceTag).strName + " {" + strSalida + "}")
  End If
End Function

Public Sub writeStyleCSS(ByVal strDirectorioDocumentacion As String)
'--> Escribe el fichero de estilos
Dim lngFichero As Long

  lngFichero = FreeFile
  Open strDirectorioDocumentacion + "Style.css" For Output As lngFichero
    Print #lngFichero, "/* Documentador HTML style sheet */"
    Print #lngFichero, "/* Define colores, fuentes y atributos de estilo  */"
    
    Print #lngFichero, "/* Background de página */"
    Print #lngFichero, getTextStyle(eTagBody)
    
    Print #lngFichero, "/* Colores de los enlaces */"
    Print #lngFichero, getTextStyle(eTagA)
    Print #lngFichero, getTextStyle(eTagALink)
    Print #lngFichero, getTextStyle(eTagAActive)
    Print #lngFichero, getTextStyle(eTagAVisited)
    
    Print #lngFichero, "/* Colores generales */"
    Print #lngFichero, getTextStyle(eTagHeader)
    Print #lngFichero, getTextStyle(eTagSubroutineHeader)
       
    Print #lngFichero, "/* Colores de tabla */"
    Print #lngFichero, getTextStyle(eTagTableCaption)
    Print #lngFichero, getTextStyle(eTagTableSubCaption)
    Print #lngFichero, getTextStyle(eTagTableRowColor)
    Print #lngFichero, getTextStyle(eTagTableRowColor)
    
    Print #lngFichero, "/* Barra de navegación */"
    Print #lngFichero, getTextStyle(eTagBotonEnabled)
    Print #lngFichero, getTextStyle(eTagBotonDisabled)
  Close lngFichero
End Sub

Public Sub seePreview(ByVal wbExplorer As WebBrowser)
'--> Muestra un preview con los estilos
'  With wbExplorer.Document
'    .documentelement.outerhtml = "<HTML> <HEAD> <TITLE> Preview </TITLE> </HEAD> <BODY> <P> Hola </BODY> </HTML>"
'  End With
End Sub

Private Sub Class_Initialize()
  ReDim arrTagEstilos(12)
  arrTagEstilos(0).uTagNombre = eTagBody
  arrTagEstilos(0).strName = "Body"
  arrTagEstilos(1).uTagNombre = eTagA
  arrTagEstilos(1).strName = "A"
  arrTagEstilos(2).uTagNombre = eTagALink
  arrTagEstilos(2).strName = "A:Link"
  arrTagEstilos(3).uTagNombre = eTagAActive
  arrTagEstilos(3).strName = "A:Active"
  arrTagEstilos(4).uTagNombre = eTagAVisited
  arrTagEstilos(4).strName = "A:Visited"
  arrTagEstilos(5).uTagNombre = eTagHeader
  arrTagEstilos(5).strName = "Header"
  arrTagEstilos(6).uTagNombre = eTagSubroutineHeader
  arrTagEstilos(6).strName = "Header Subroutine"
  arrTagEstilos(7).uTagNombre = eTagTableCaption
  arrTagEstilos(7).strName = "Table Caption"
  arrTagEstilos(8).uTagNombre = eTagTableSubCaption
  arrTagEstilos(8).strName = "Table subcaption"
  arrTagEstilos(9).uTagNombre = eTagTableRowColor
  arrTagEstilos(9).strName = "Color fila"
  arrTagEstilos(10).uTagNombre = eTagRowHeader
  arrTagEstilos(10).strName = "Color de Cabecera"
  arrTagEstilos(11).uTagNombre = eTagBotonEnabled
  arrTagEstilos(11).strName = "Botón activo"
  arrTagEstilos(12).uTagNombre = eTagBotonDisabled
  arrTagEstilos(12).strName = "Botón inactivo"
End Sub
