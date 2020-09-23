VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase para almacenar las constantes, variables, contenidos de tipos ...
Option Explicit

Private Enum eTipoVariable 'Tipo de la variable
  etvConstante = 0 'Constante
  etvVariable 'Variable
End Enum

Private Type tVariable 'Tipo para almacenar los parámetros de la variables
  TipoVariable As eTipoVariable 'Tipo
  Ambito As etAmbito 'Ambito
  Estatica As Boolean 'Estática o no
  Nombre As String 'Nombre de variable - constante
  Tipo As String 'Tipo
  Valor As String 'Valor (en el caso de constantes)
  Comentario As String 'Comentario
End Type

Private TablaVariables() As tVariable 'Tabla para almacenar las variables

Public Sub Add(ByVal Cadena As String)
'--> Parte la cadena en trozos para encontrar sus parámetros y la guarda sobre TablaVariables
Dim CadIntermedia As String, CadInicial As String
Dim AmbitoAnterior As etAmbito
Dim TipoVariableAnterior As eTipoVariable

  'CadInicial = guObjGeneral.QuitarParametro(Cadena, ",")
  TablaVariables(UBound(TablaVariables)).Comentario = Cadena
  Cadena = guObjGeneral.QuitarParametro(TablaVariables(UBound(TablaVariables)).Comentario, "'")
  CadInicial = guObjGeneral.QuitarParametro(Cadena, ",")
  AmbitoAnterior = etAmbPrivado
  TipoVariableAnterior = etvVariable
  While CadInicial <> ""
    With TablaVariables(UBound(TablaVariables))
      .Ambito = AmbitoAnterior
      .TipoVariable = TipoVariableAnterior
      .Estatica = False
      Do
        CadIntermedia = guObjGeneral.QuitarParametro(CadInicial, " ")
        If CadIntermedia <> "" Then
          Select Case UCase$(CadIntermedia)
            'Ambito de la variable
            Case "PRIVATE", "DIM"
              .Ambito = etAmbPrivado
            Case "PUBLIC"
              .Ambito = etAmbPublico
            Case "GLOBAL"
              .Ambito = etAmbGlobal
            Case "PROTECTED"
              .Ambito = etAmbProtegido
            'Tipo variable
            Case "CONST"
              .TipoVariable = etvConstante
            'Estática
            Case "STATIC"
              .Estatica = True
            'Tipo
            Case "AS"
              CadIntermedia = guObjGeneral.QuitarParametro(CadInicial, " ")
              If UCase$(CadIntermedia) = "NEW" Then CadIntermedia = guObjGeneral.QuitarParametro(CadInicial, " ")
              .Tipo = CadIntermedia
            'En el caso de las constantes tienen un = valor
            Case "="
              CadIntermedia = guObjGeneral.QuitarParametro(CadInicial, "'")
              CadInicial = "" 'Para los casos en que el valor sea una cadinicial
              .Valor = CadIntermedia
            'Nombre
            Case Else
              .Nombre = CadIntermedia
          End Select
        End If
      Loop Until CadInicial = ""
      AmbitoAnterior = .Ambito
      TipoVariableAnterior = .TipoVariable
    End With
    ReDim Preserve TablaVariables(UBound(TablaVariables) + 1)
    CadInicial = guObjGeneral.QuitarParametro(Cadena, ",")
  Wend
End Sub

Private Function getHTML(ByRef clGeneradorHTML As ClassHTML, ByVal Tipo As eTipoVariable) As String
'--> Obtiene una cadena HTML para las variables de tipo <B> Tipo </B>
Dim Indice As Integer
Dim ExistenValores As Boolean

  getHTML = "<TR>"
  getHTML = getHTML + "<TD CLASS=TableSubcaption> Nombre </TD>"
  getHTML = getHTML + "<TD CLASS=TableSubcaption> Ambito </TD>"
  getHTML = getHTML + "<TD CLASS=TableSubcaption> Estática </TD>"
  getHTML = getHTML + "<TD CLASS=TableSubcaption> Tipo </TD>"
  If Tipo = etvConstante Then getHTML = getHTML + "<TD CLASS=TableSubcaption> Valor </TD>"
  getHTML = getHTML + "<TD CLASS=TableSubcaption> Comentario </TD>"
  getHTML = getHTML + " </TR>"
  ExistenValores = False
  For Indice = 0 To UBound(TablaVariables) - 1
    With TablaVariables(Indice)
      If .TipoVariable = Tipo Then
        ExistenValores = True
        getHTML = getHTML + "<TR>"
        getHTML = getHTML + "<TD>" + clGeneradorHTML.getCadenaNegrita(clGeneradorHTML.getCadenaConColor("0000FF", .Nombre)) + " </TD>"
        getHTML = getHTML + "<TD>" + getCadenaAmbito(.Ambito) + " </TD>"
        getHTML = getHTML + "<TD>" + IIf(.Estatica, "Sí", "No") + " </TD>"
        getHTML = getHTML + "<TD>" + clGeneradorHTML.getCadenaVacia(.Tipo) + " </TD>"
        If Tipo = etvConstante Then getHTML = getHTML + "<TD>" + clGeneradorHTML.getCadenaVacia(.Valor) + " </TD>"
        getHTML = getHTML + "<TD>" + clGeneradorHTML.getCadenaVacia(.Comentario) + " </TD>"
        getHTML = getHTML + " </TR>"
      End If
    End With
  Next Indice
  getHTML = IIf(ExistenValores, clGeneradorHTML.getTable(IIf(Tipo = etvVariable, "Variables", "Constantes"), getHTML, False), "")
End Function

Public Function getHTMLTipo(ByRef clGeneradorHTML As ClassHTML) As String
'--> Obtiene un código HTML útil para meterlo en los tipos o enumerados
Dim Indice As Integer

  getHTMLTipo = ""
  For Indice = 0 To UBound(TablaVariables) - 1
    With TablaVariables(Indice)
      getHTMLTipo = getHTMLTipo + "<TR>"
      getHTMLTipo = getHTMLTipo + "<TD>" + clGeneradorHTML.getCadenaNegrita(.Nombre) + " </TD>"
      getHTMLTipo = getHTMLTipo + "<TD>" + clGeneradorHTML.getCadenaVacia(IIf(.Tipo = "", .Valor, .Tipo)) + " </TD>"
      getHTMLTipo = getHTMLTipo + "<TD>" + clGeneradorHTML.getCadenaVacia(.Comentario) + " </TD>"
      getHTMLTipo = getHTMLTipo + "</TR>"
    End With
  Next Indice
End Function

Public Sub writeHTML(ByRef clGeneradorHTML As ClassHTML, ByVal NumeroFichero As Long)
'--> Escribe las variables del formulario sobre el fichero HTML
'--> @sub getHTML
Dim CadSalida As String

  CadSalida = getHTML(clGeneradorHTML, etvConstante)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etvVariable)
  If CadSalida <> "" Then Print #NumeroFichero, clGeneradorHTML.getAnchor("Variables") + CadSalida
End Sub

Private Sub Class_Initialize()
  ReDim TablaVariables(0)
End Sub

Private Sub Class_Terminate()
  Erase TablaVariables
End Sub
