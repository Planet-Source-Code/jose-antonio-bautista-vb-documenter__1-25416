VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase para almacenar los tipos y enumerados
Option Explicit

Private Enum etTipos 'Enumerado para el tipo
  etTipo = 0 'Es un tipo
  etEnum 'Es un enumerado
End Enum

Private Type tTipos 'Tipo con los parámetros del tipo
  Tipo As etTipos 'Tipo
  Ambito As etAmbito 'Ambito
  Nombre As String 'Nombre
  Comentario As String 'Comentario
  clVariables As ClassVariables 'Clase con las variables (parte del tipo o del enumerado)
End Type

Private TablaTipos() As tTipos 'Tabla para almacenar todos los tipos del formulario

Public Sub addVariable(ByVal Cadena As String)
'--> Añade las variables al tipo
  TablaTipos(UBound(TablaTipos) - 1).clVariables.Add Cadena
End Sub

Public Sub Add(ByVal Cadena As String)
'--> Añade un tipo a la tabla de tipos
Dim CadIntermedia As String

  With TablaTipos(UBound(TablaTipos))
    Set .clVariables = New ClassVariables
    .Comentario = Cadena
    Cadena = guObjGeneral.QuitarParametro(.Comentario, "'")
    While Cadena <> ""
      CadIntermedia = guObjGeneral.QuitarParametro(Cadena, " ")
      .Ambito = etAmbPublico
      Select Case UCase$(CadIntermedia)
        Case CadIntermedia
          'Ambito
          Case "PRIVATE"
            .Ambito = etAmbPrivado
          Case "PUBLIC"
            .Ambito = etAmbPublico
          Case "PROTECTED"
            .Ambito = etAmbProtegido
          Case "GLOBAL"
            .Ambito = etAmbGlobal
          'Tipo
          Case "TYPE"
            .Tipo = etTipo
          Case "ENUM"
            .Tipo = etEnum
          'Nombre
          Case Else
            .Nombre = CadIntermedia
      End Select
    Wend
  End With
  ReDim Preserve TablaTipos(UBound(TablaTipos) + 1)
End Sub

Private Function getHTML(ByRef clGeneradorHTML As ClassHTML, ByVal Tipo As etTipos) As String
'--> Obtiene una cadena HTML para las variables de tipo <B> Tipo </B>
Dim Indice As Integer
Dim blnExistenValores As Boolean
Dim strCadTipo As String

  getHTML = "<TD CLASS=TableSubcaption>" + clGeneradorHTML.getCadenaNegrita("Nombre")
  getHTML = getHTML + "<TD CLASS=TableSubcaption> Ambito </TD>"
  getHTML = getHTML + "<TD CLASS=TableSubcaption> Comentario </TD>"
  getHTML = getHTML + "<TD CLASS=TableSubcaption>" + IIf(Tipo = etEnum, "Valores", "Miembros") + " </TD>"
  blnExistenValores = False
  For Indice = 0 To UBound(TablaTipos) - 1
    With TablaTipos(Indice)
      If .Tipo = Tipo Then
        blnExistenValores = True
        getHTML = getHTML + "<TR>"
        getHTML = getHTML + "<TD>" + clGeneradorHTML.getCadenaNegrita(clGeneradorHTML.getCadenaConColor("0000FF", .Nombre)) + " </TD>"
        getHTML = getHTML + "<TD>" + getCadenaAmbito(.Ambito) + " </TD>"
        getHTML = getHTML + "<TD>" + clGeneradorHTML.getCadenaVacia(.Comentario) + " </TD>"
        strCadTipo = .clVariables.getHTMLTipo(clGeneradorHTML)
        If strCadTipo <> "" Then
          getHTML = getHTML + "<TD> " + clGeneradorHTML.getTable("", _
                                                   "<TD CLASS=TableSubcaption>" + clGeneradorHTML.getCadenaNegrita("Nombre") + _
                                                   "<TD CLASS=TableSubcaption>" + clGeneradorHTML.getCadenaNegrita(IIf(.Tipo = etEnum, "Valor", "Tipo")) + _
                                                   "<TD CLASS=TableSubcaption>" + clGeneradorHTML.getCadenaNegrita("Comentario") + _
                                                   strCadTipo, False) + " </TD>"
        End If
        getHTML = getHTML + " </TR>"
      End If
    End With
  Next Indice
  getHTML = IIf(blnExistenValores, clGeneradorHTML.getTable(IIf(Tipo = etEnum, "Enumerados", "Tipos"), getHTML), "")
End Function

Public Sub writeHTML(ByRef clGeneradorHTML As ClassHTML, ByVal NumeroFichero As Long)
'--> Escribe los tipos y enumerados del formulario sobre el fichero HTML
'--> @sub getHTML
Dim CadSalida As String

  CadSalida = getHTML(clGeneradorHTML, etEnum)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etTipo)
  If CadSalida <> "" Then Print #NumeroFichero, CadSalida
End Sub

Private Sub Class_Initialize()
  ReDim TablaTipos(0)
End Sub

Private Sub Class_Terminate()
Dim Indice As Integer

  For Indice = 0 To UBound(TablaTipos) - 1
    Set TablaTipos(Indice).clVariables = Nothing
  Next Indice
  Erase TablaTipos
End Sub
