VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassComentarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase de comentarios
'--> Lista de tags:
'--> <UL>
'--> <LI> <B> @author </B> Nombre del autor </LI>
'--> <LI> <B> @see </B> Ver otros documentos (hiperenlace) </LI>
'--> <LI> <B> @version </B> Versión </LI>
'--> <LI> <B> @revision </B> Datos de revisión </LI>
'--> <LI> <B> @date </B> Fecha </LI>
'--> <LI> <B> @sub </B> Subrutinas a las que llama </LI>
'--> <LI> <B> @form </B> Formularios a los que llama </LI>
'--> <LI> <B> @deprecated </B> Se utiliza la versión nueva </LI>
'--> <LI> <B> @exception </B> Descripción de excepciones </LI>
'--> <LI> <B> @param </B> Descripción de parámetro </LI>
'--> <LI> <B> @return </B> Valor de retorno </LI>
'--> </UL>
Option Explicit

Private Enum etComentario 'Enumerado con los comentarios posibles
  etComNormal = 0 'Normal
  etComAuthor 'Autor
  etComDeprecated 'Versión no utilizada
  etComException 'Excepción
  etComParam 'Parámetro de subrutina / función
  etComReturn 'Valor de retorno
  etComSee 'Ver también
  etComVersion 'Versión
  etComDate 'Fecha
  etComSub 'Subrutina a la que llama
  etComForm 'Formulario
  etComRevision 'Número de revisión
End Enum

Private Type tComentario 'Tipo para comentarios
  Tipo As etComentario 'Tipo del comentario
  Comentario As String 'Cadena del comentario en sí
End Type

Private TablaComentarios() As tComentario 'Tabla con todos los comentarios de esa clase

Public Sub Add(ByVal Cadena As String)
'--> Añade un comentario a la tabla de comentarios
Dim CadTipo As String

  With TablaComentarios(UBound(TablaComentarios))
    Cadena = Trim(Mid$(Cadena, 5, Len(Cadena)))
    CadTipo = guObjGeneral.QuitarParametro(Cadena, " ")
    Select Case UCase$(CadTipo)
      Case "@AUTHOR"
        .Tipo = etComAuthor
      Case "@DEPRECATED"
        .Tipo = etComDeprecated
      Case "@EXCEPTION"
        .Tipo = etComException
      Case "@PARAM"
        .Tipo = etComParam
      Case "@RETURN"
        .Tipo = etComReturn
      Case "@SEE"
        .Tipo = etComSee
      Case "@VERSION"
        .Tipo = etComVersion
      Case "@DATE"
        .Tipo = etComDate
      Case "@SUB"
        .Tipo = etComSub
      Case "@FORM"
        .Tipo = etComForm
      Case "@REVISION"
        .Tipo = etComRevision
      Case Else
        .Tipo = etComNormal
    End Select
    If .Tipo = etComNormal Then
      Cadena = CadTipo + " " + Cadena
    ElseIf .Tipo = etComParam Or .Tipo = etComSub Or .Tipo = etComForm Then
      Cadena = "<B> " + guObjGeneral.QuitarParametro(Cadena, " ") + " </B> " + Cadena
    End If
    .Comentario = Cadena
  End With
  ReDim Preserve TablaComentarios(UBound(TablaComentarios) + 1)
End Sub

Private Function getHTML(ByRef clGeneradorHTML As ClassHTML, ByVal Tipo As etComentario, ByVal blnDesdeFormulario As Boolean) As String
'--> Obtiene una cadena HTML para los comentarios de tipo <B> Tipo </B>
Dim Indice As Integer
Dim ExistenValores As Boolean, PorPartes As Boolean

  PorPartes = False
  Select Case Tipo
    Case etComNormal
      getHTML = IIf(blnDesdeFormulario, "", "<LI>" + clGeneradorHTML.getCadenaConColor("0000FF", clGeneradorHTML.getCadenaNegrita("Comentarios: ")))
      PorPartes = True
    Case etComAuthor
      getHTML = "<LI > " + clGeneradorHTML.getCadenaConColor("0000FF", clGeneradorHTML.getCadenaNegrita("Autor: "))
    Case etComDeprecated
      getHTML = "<LI>" + clGeneradorHTML.getCadenaConColor("0000FF", clGeneradorHTML.getCadenaNegrita("Deprecated: "))
    Case etComException
      getHTML = "<LI>" + clGeneradorHTML.getCadenaConColor("0000FF", clGeneradorHTML.getCadenaNegrita("Excepción: "))
      PorPartes = True
    Case etComParam
      getHTML = "<LI>" + clGeneradorHTML.getCadenaConColor("0000FF", clGeneradorHTML.getCadenaNegrita("Parámetros: "))
      PorPartes = True
    Case etComReturn
      getHTML = "<LI>" + clGeneradorHTML.getCadenaConColor("0000FF", clGeneradorHTML.getCadenaNegrita("Retorno: "))
      PorPartes = True
    Case etComSee
      getHTML = "<LI>" + clGeneradorHTML.getCadenaConColor("0000FF", clGeneradorHTML.getCadenaNegrita("See: "))
      PorPartes = True
    Case etComVersion
      getHTML = "<LI>" + clGeneradorHTML.getCadenaConColor("0000FF", clGeneradorHTML.getCadenaNegrita("Versión: "))
    Case etComDate
      getHTML = "<LI>" + clGeneradorHTML.getCadenaConColor("0000FF", clGeneradorHTML.getCadenaNegrita("Fecha: "))
    Case etComRevision
      getHTML = "<LI>" + clGeneradorHTML.getCadenaConColor("0000FF", clGeneradorHTML.getCadenaNegrita("Revisión: "))
      PorPartes = True
    Case etComSub
      getHTML = "<LI>" + clGeneradorHTML.getCadenaConColor("0000FF", clGeneradorHTML.getCadenaNegrita("Rutinas: "))
      PorPartes = True
    Case etComForm
      getHTML = "<LI>" + clGeneradorHTML.getCadenaConColor("0000FF", clGeneradorHTML.getCadenaNegrita("Formularios: "))
      PorPartes = True
  End Select
  If PorPartes And Not blnDesdeFormulario Then getHTML = getHTML + "<UL>"
  ExistenValores = False
  For Indice = 0 To UBound(TablaComentarios) - 1
    With TablaComentarios(Indice)
      If .Tipo = Tipo Then
        ExistenValores = True
        getHTML = getHTML + IIf(PorPartes And Not blnDesdeFormulario, "<LI>", "<P>") + .Comentario
      End If
    End With
  Next Indice
  If PorPartes And Not blnDesdeFormulario Then
    getHTML = getHTML + "</UL>"
  End If
  If Not ExistenValores Then getHTML = ""
End Function

Public Sub writeHTML(ByRef clGeneradorHTML As ClassHTML, ByVal NumeroFichero As Long)
'--> Escribe los comentarios de la rutina sobre el fichero HTML
'--> @sub getHTML
Dim CadSalida As String, CadIntermedia As String

  CadSalida = getHTML(clGeneradorHTML, etComNormal, False)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComAuthor, False)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComDate, False)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComVersion, False)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComRevision, False)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComDeprecated, False)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComParam, False)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComReturn, False)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComForm, False)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComSub, False)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComException, False)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComSee, False)
  If CadSalida <> "" Then Print #NumeroFichero, "<UL>" + CadSalida + "</UL>" '<HR>"
End Sub

Public Function getStrHTML(ByRef clGeneradorHTML As ClassHTML) As String
'--> Obtiene una cadena HTML con los comentarios de la cabecera del formulario para meterlo en la tabla
'--> @sub getHTML
Dim CadSalida As String, CadIntermedia As String

  CadSalida = getHTML(clGeneradorHTML, etComAuthor, True)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComDate, True)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComVersion, True)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComRevision, True)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComDeprecated, True)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComParam, True)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComReturn, True)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComForm, True)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComSub, True)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComException, True)
  CadSalida = CadSalida + getHTML(clGeneradorHTML, etComSee, True)
  getStrHTML = CadSalida + getHTML(clGeneradorHTML, etComNormal, True)
End Function

Private Sub Class_Initialize()
  ReDim TablaComentarios(0)
End Sub

Private Sub Class_Terminate()
  Erase TablaComentarios
End Sub
