VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFormulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase de definición de formularios / ficheros integrantes de un proyecto VB
Option Explicit

Public Enum eTipo 'Tipos posibles de ficheros
  eTipoFormulario = 0 'Formulario
  eTipoModulo 'Módulo BAS
  eTipoClase 'Módulo de clase
  eTipoControlUsuario 'Control de usuario
  eTipoDocumentoUsuario 'Documento de usuario
End Enum

Public Enum eTipoLinea 'Tipo de la línea que se está leyendo
  eTLConstante = 0
  eTLVariable
  eTLTipo
  eTLEnumerado
  eTLEvento
  eTLDeclaracion
  eTLRutina
  eTLControl
  eTLComentario
  eTLFinBloque
  eTLLineaInstruccion
End Enum

Private etTipoFormulario As eTipo
Private strNombreFichero As String
Private strNombreFormulario As String

Private clControl As New ClassControl 'Clase con los controles de este formulario
Private clVariables As New ClassVariables 'Clase con las variables de este formulario
Private clComentarios As New ClassComentarios 'Clase con los comentarios iniciales
Private clRutinas As New ClassRutina 'Clase con las rutinas
Private clTipos As New ClassTipos 'Clase con los tipos / enumerados

Public Property Let Tipo(ByVal petTipo As eTipo)
  etTipoFormulario = petTipo
End Property

Public Property Get Tipo() As eTipo
  Tipo = etTipoFormulario
End Property

Public Property Let NombreFichero(ByVal pstrNombreFichero As String)
  strNombreFichero = pstrNombreFichero
End Property

Public Property Get NombreFichero() As String
  NombreFichero = strNombreFichero
End Property

Public Property Let NombreFormulario(ByVal pstrNombreFormulario As String)
'--> Establece el nombre del fichero (le quita el directorio)
Dim Indice As Integer
Dim Encontrado As Boolean

  strNombreFormulario = pstrNombreFormulario
  Indice = Len(strNombreFormulario)
  Encontrado = False
  While Not Encontrado And Indice > 0
    If Mid$(strNombreFormulario, Indice, 1) = "\" Then
      Encontrado = True
    Else
      Indice = Indice - 1
    End If
  Wend
  If Encontrado Then strNombreFormulario = Mid$(strNombreFormulario, Indice + 1, Len(strNombreFormulario))
  strNombreFormulario = guObjGeneral.QuitarParametro(strNombreFormulario, ".")
End Property

Public Property Get NombreFormulario() As String
  NombreFormulario = strNombreFormulario
End Property

Private Function LeerLinea(ByVal NumeroFichero As Integer) As String
'--> Lee una o varias líneas del fichero a documentar, las junta si están separada por "_"
Dim Cadena As String, SegundaCadena As String

  On Error GoTo ErrorLectura
  Cadena = ""
  If Not EOF(NumeroFichero) Then
    Line Input #NumeroFichero, Cadena
    While Right$(Cadena, 1) = "_" And Not EOF(NumeroFichero)
      Line Input #NumeroFichero, SegundaCadena
      Cadena = Left$(Cadena, Len(Cadena) - 1) + SegundaCadena
    Wend
  End If
  
ErrorLectura:
  LeerLinea = Cadena 'Siempre voy a devolver la cadena que tenga en ese momento
End Function

Private Function RecogerTipoLinea(ByVal CadEntrada As String) As eTipoLinea
'--> Coge la cadena leída del fichero e intenta dilucidar que es mirando la primera palabra
'--> @return Tipo de línea leída
Dim Token As String, TokenSecundario As String

  CadEntrada = UCase$(CadEntrada)
  Token = guObjGeneral.QuitarParametro(CadEntrada, " ")
  If Token = "PRIVATE" Or Token = "PUBLIC" Or Token = "GLOBAL" Then
    TokenSecundario = guObjGeneral.QuitarParametro(CadEntrada, " ")
    If TokenSecundario = "SUB" Or TokenSecundario = "FUNCTION" Or TokenSecundario = "PROPERTY" Or _
        TokenSecundario = "EVENT" Or TokenSecundario = "DECLARE" Or TokenSecundario = "CONST" Or _
        TokenSecundario = "TYPE" Or TokenSecundario = "ENUM" Then
      'Así nos liberamos de las variables privadas, públicas... que no tienen la forma "private dim"
      Token = TokenSecundario
    Else
      Token = "DIM" 'Pequeño truco para las variables globales
    End If
  End If
  Select Case Token
    Case "CONST"
      RecogerTipoLinea = eTLConstante
    Case "DIM", "STATIC"
      RecogerTipoLinea = eTLVariable
    Case "TYPE"
      RecogerTipoLinea = eTLTipo
    Case "ENUM"
      RecogerTipoLinea = eTLEnumerado
    Case "EVENT"
      RecogerTipoLinea = eTLEvento
    Case "DECLARE"
      RecogerTipoLinea = eTLDeclaracion
    Case "SUB", "FUNCTION", "PROPERTY"
      RecogerTipoLinea = eTLRutina
    Case "BEGIN", "CONTROL"
      RecogerTipoLinea = eTLControl
    Case "'-->"
      RecogerTipoLinea = eTLComentario
    Case "END"
      TokenSecundario = guObjGeneral.QuitarParametro(CadEntrada, " ")
      If TokenSecundario = "SUB" Or TokenSecundario = "FUNCTION" Or TokenSecundario = "PROPERTY" Or TokenSecundario = "TYPE" Or TokenSecundario = "ENUM" Then
        RecogerTipoLinea = eTLFinBloque
      Else
        RecogerTipoLinea = eTLLineaInstruccion
      End If
    Case Else
      RecogerTipoLinea = eTLLineaInstruccion
  End Select
End Function

Private Function getNombreFichero(ByVal Directorio As String, ByVal NombreFichero As String) As String
'--> Obtiene el nombre del fichero, útil en los casos que el nombre de fichero sea del tipo ..\ModuloDatos\FR_Datos.frm
  If Right$(Directorio, 1) = "\" Then Directorio = Left$(Directorio, Len(Directorio) - 1)
  While Left$(NombreFichero, 3) = "..\"
    'Quita el último directorio a la cadena directorio
    NombreFichero = Mid$(NombreFichero, 4, Len(NombreFichero))
    While Right$(Directorio, 1) <> "\" And Directorio <> ""
      Directorio = Left$(Directorio, Len(Directorio) - 1)
    Wend
  Wend
  If Right$(Directorio, 1) = "\" Then Directorio = Left$(Directorio, Len(Directorio) - 1)
  getNombreFichero = Directorio + "\" + NombreFichero
End Function

Public Sub read(ByVal Directorio As String)
'--> Lee el formulario (el fichero), va leyendo cada una de sus líneas y pasando a las distintas clases
'--> @sub LeerLinea
Dim NumeroFichero As Long
Dim CadenaLectura As String
Dim isSubrutina As Boolean, isTipo As Boolean

  On Error GoTo ErrorLectura
  NumeroFichero = FreeFile
  Open getNombreFichero(Trim(Directorio), Trim(strNombreFichero)) For Input As NumeroFichero
    isSubrutina = False
    isTipo = False
    CadenaLectura = LeerLinea(NumeroFichero)
    While Not EOF(NumeroFichero)
      Select Case RecogerTipoLinea(CadenaLectura)
        Case eTLControl
          clControl.addControl CadenaLectura
        Case eTLConstante
          clVariables.Add CadenaLectura
        Case eTLVariable
          If isTipo Then
            clTipos.addVariable CadenaLectura
          ElseIf isSubrutina Then
            clRutinas.addVariable CadenaLectura
          Else
            clVariables.Add CadenaLectura
          End If
        Case eTLTipo
          isTipo = True
          clTipos.Add CadenaLectura
        Case eTLEnumerado
          isTipo = True
          clTipos.Add CadenaLectura
        Case eTLEvento
          clRutinas.Add CadenaLectura
        Case eTLDeclaracion
          clRutinas.Add CadenaLectura
        Case eTLRutina
          isSubrutina = True
          clRutinas.Add CadenaLectura
        Case eTLComentario
          If isTipo Then
          ElseIf isSubrutina Then
            clRutinas.addComentarios CadenaLectura
          Else
            clComentarios.Add CadenaLectura
          End If
        Case eTLFinBloque
          If isTipo Then
            isTipo = False
          ElseIf isSubrutina Then
            isSubrutina = False
          End If
        Case eTLLineaInstruccion
          If isTipo Then clTipos.addVariable CadenaLectura
      End Select
      CadenaLectura = LeerLinea(NumeroFichero)
    Wend
  Close NumeroFichero
  Exit Sub
  
ErrorLectura:
End Sub

Public Sub writeHTML(ByVal Directorio As String, ByVal blnConListaParametros As Boolean, _
                     ByVal blnConVariables As Boolean, ByVal strProyecto As String, ByRef clGeneradorHTML As ClassHTML)
'--> Escribe el fichero HTML del formulario documentado (el frame de la derecha), _
     va pasando a las diferentes clases para que escriban su parte
Dim NumeroFichero As Long
Dim strTablaCabecera As String

  NumeroFichero = FreeFile
  Open Directorio + "\" + strNombreFormulario + ".html" For Output As NumeroFichero
    'Graba la cabecera
    Print #NumeroFichero, clGeneradorHTML.getCabeceraHTML("Documentación " + strNombreFormulario)
    'Graba la barra de navegador al principio
    Print #NumeroFichero, clGeneradorHTML.getBarraNavegadorHTML(True, True, _
                    (Tipo = eTipoFormulario Or Tipo = eTipoControlUsuario Or Tipo = eTipoDocumentoUsuario), _
                    True, True)
    'Graba la descripción del formulario
    Print #NumeroFichero, clGeneradorHTML.getAnchor("Descripcion")
    strTablaCabecera = "<TR> <TD CLASS=TableCaption> Proyecto </TD> <TD> " + strProyecto + "</TD></TR>"
    strTablaCabecera = strTablaCabecera + "<TR> <TD CLASS=TableCaption> Nombre </TD> <TD> " + strNombreFormulario + "</TD></TR>"
    strTablaCabecera = strTablaCabecera + "<TR> <TD CLASS=TableCaption> Comentarios </TD> <TD> " + _
                                  clComentarios.getStrHTML(clGeneradorHTML) + "</TD></TR>"
    Print #NumeroFichero, clGeneradorHTML.getTable("Descripción", strTablaCabecera, False)
    Print #NumeroFichero, "<HR>"
    'Graba todas sus partes
    clControl.writeHTML clGeneradorHTML, NumeroFichero
    clTipos.writeHTML clGeneradorHTML, NumeroFichero
    clVariables.writeHTML clGeneradorHTML, NumeroFichero
    clRutinas.writeHTML clGeneradorHTML, blnConListaParametros, blnConVariables, NumeroFichero
    'Obtiene la barra de navegador y el cierre
    Print #NumeroFichero, clGeneradorHTML.getBarraNavegadorHTML(True, True, _
                    (Tipo = eTipoFormulario Or Tipo = eTipoControlUsuario Or Tipo = eTipoDocumentoUsuario), _
                    True, True)
    Print #NumeroFichero, clGeneradorHTML.getCierreHTML()
  Close NumeroFichero
End Sub

Private Sub Class_Terminate()
  Set clControl = Nothing
  Set clVariables = Nothing
  Set clComentarios = Nothing
  Set clRutinas = Nothing
  Set clTipos = Nothing
End Sub
