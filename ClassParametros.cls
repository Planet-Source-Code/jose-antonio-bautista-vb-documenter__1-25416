VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase de parámetros de una rutina
Option Explicit

Private Enum eTipoParametro 'Tipo de los parámetros
  etParPorValor = 0 'Parámetro por valor
  etParPorReferencia 'Parámetro por referencia
End Enum

Private Type tParametro 'Tipo para almacenar los parámetros
  Tipo As eTipoParametro 'Tipo del parámetro (valor o referencia)
  Nombre As String 'Nombre de parámetro
  TipoVariable As String 'Tipo variable (integer, double, ...)
  Opcional As Boolean 'Indica si es opcional o no
  Valor As String 'Si es opcional su valor por defecto
End Type

Private TablaParametros() As tParametro 'Tabla con los parámetros de la rutina

Public Sub Clear()
'--> Limpia la tabla
  ReDim TablaParametros(0)
End Sub

Private Sub addParametro(ByVal Cadena As String)
'--> Añade un parámetro a la tabla
Dim CadIntermedia As String

  With TablaParametros(UBound(TablaParametros))
    .Tipo = etParPorReferencia
    While Cadena <> ""
      CadIntermedia = guObjGeneral.QuitarParametro(Cadena, " ")
      If CadIntermedia <> "" Then
        Select Case UCase$(CadIntermedia)
          Case "BYVAL"
            .Tipo = etParPorValor
          Case "BYREF"
            .Tipo = etParPorReferencia
          Case "OPTIONAL"
            .Opcional = True
          Case "="
            .Valor = Cadena
            Cadena = ""
          Case "AS"
            .TipoVariable = guObjGeneral.QuitarParametro(Cadena, " ")
          Case Else
            .Nombre = CadIntermedia
        End Select
      End If
    Wend
  End With
  ReDim Preserve TablaParametros(UBound(TablaParametros) + 1)
End Sub

Public Sub Add(ByVal Cadena As String)
'--> Añade los parámetros que estén en la cadena a la tabla parámetros
  Cadena = Trim(Cadena)
  If Left$(Cadena, 1) = "(" Then Cadena = Right$(Cadena, Len(Cadena) - 1)
  If Right$(Cadena, 1) = ")" Then Cadena = Left$(Cadena, Len(Cadena) - 1)
  While Cadena <> ""
    addParametro Trim(guObjGeneral.QuitarParametro(Cadena, ","))
  Wend
End Sub

Public Sub writeHTML(ByRef clGeneradorHTML As ClassHTML, ByVal NumeroFichero As Long)
'--> Escribe los parámetros de la rutina sobre el fichero HTML
Dim Indice As Integer
Dim CadSalida As String

  If UBound(TablaParametros) > 0 Then
    CadSalida = CadSalida + "<TR>"
    CadSalida = CadSalida + "<TD CLASS=TableSubcaption> Nombre"
    CadSalida = CadSalida + "<TD CLASS=TableSubcaption> Tipo Parámetro"
    CadSalida = CadSalida + "<TD CLASS=TableSubcaption> Opcional"
    CadSalida = CadSalida + "<TD CLASS=TableSubcaption> Tipo"
    CadSalida = CadSalida + "<TD CLASS=TableSubcaption> Valor"
    For Indice = 0 To UBound(TablaParametros) - 1
      With TablaParametros(Indice)
        CadSalida = CadSalida + "<TR>"
        CadSalida = CadSalida + "<TD>" + clGeneradorHTML.getCadenaNegrita(clGeneradorHTML.getCadenaConColor("0000FF", .Nombre))
        CadSalida = CadSalida + "<TD>" + IIf(.Tipo = etParPorValor, "Por Valor", "Por Referencia")
        CadSalida = CadSalida + "<TD>" + IIf(.Opcional, "Sí", "No")
        CadSalida = CadSalida + "<TD>" + clGeneradorHTML.getCadenaVacia(.TipoVariable)
        CadSalida = CadSalida + "<TD>" + clGeneradorHTML.getCadenaVacia(.Valor)
      End With
    Next Indice
    Print #NumeroFichero, clGeneradorHTML.getTable("Parámetros", CadSalida, False)
  End If
End Sub

Private Sub Class_Initialize()
  ReDim TablaParametros(0)
End Sub

Private Sub Class_Terminate()
  Erase TablaParametros
End Sub
