VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRutina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase para guardar subrutinas / funciones / propiedades
Option Explicit

Private Enum etRutina 'Enumerado con los tipos de rutinas posibles
  etRutSubRutina = 0 'Subrutina
  etRutFuncion 'Función
  etRutPropiedad 'Propiedad
  etRutEvento 'Evento
  etRutDeclare 'Declare
End Enum

Private Enum etPropiedad 'Enumerado con los tipos de propiedad
  etPropLet = 0 'Propiedad tipo Let
  etPropGet 'Propiedad tipo Get
  etPropSet 'Propiedad tipo Set
End Enum

Private Type tRutina 'Tipo para almacenar los valores de las rutinas
  Ambito As etAmbito 'Ambito
  Tipo As etRutina 'Tipo de rutina
  TipoPropiedad As etPropiedad 'En el caso de una propiedad su tipo
  Nombre As String 'Nombre
  CadenaCompleta As String 'Prototipo de la rutina
  TipoRetorno As String 'Tipo de retorno
  Comentarios As String 'Comentarios
  Libreria As String 'Librería
  Alias As String 'Alias
  clParametros As ClassParametros 'Clase con los parámetros
  clVariables As ClassVariables 'Clase con las variables locales
  clComentarios As ClassComentarios 'Clase con los comentarios internos a la rutina
End Type

Private TablaRutinas() As tRutina 'Tabla con las rutinas

Public Sub Add(ByVal Cadena As String)
'--> Añade una rutina a la tabla rutinas
'--> Desde la cadena leída del fichero por la clase formulario saca todos los parámetros de la rutina
Dim CadTipoNombre As String, CadParametros As String, CadFinal As String
Dim CadIntermedia As String

  CadTipoNombre = guObjGeneral.QuitarParametro(Cadena, "(")
  While Right$(Cadena, 1) <> ")"
    CadFinal = Right$(Cadena, 1) + CadFinal
    Cadena = Left$(Cadena, Len(Cadena) - 1)
  Wend
  CadFinal = Trim(CadFinal)
  CadParametros = Trim(Left$(Cadena, Len(Cadena) - 1))
  With TablaRutinas(UBound(TablaRutinas))
    Set .clParametros = New ClassParametros
    Set .clVariables = New ClassVariables
    Set .clComentarios = New ClassComentarios
    .CadenaCompleta = CadTipoNombre + "(" + CadParametros + ") " + CadFinal
    'Obtiene la parte del nombre
    .Ambito = etAmbPublico
    While CadTipoNombre <> ""
      CadIntermedia = guObjGeneral.QuitarParametro(CadTipoNombre, " ")
      Select Case UCase$(CadIntermedia)
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
        Case "SUB"
          .Tipo = etRutSubRutina
        Case "FUNCTION"
          If .Tipo <> etRutDeclare Then .Tipo = etRutFuncion
        Case "PROPERTY"
          .Tipo = etRutPropiedad
          CadIntermedia = UCase$(guObjGeneral.QuitarParametro(CadTipoNombre, " "))
          If CadIntermedia = "LET" Then
            .TipoPropiedad = etPropLet
          ElseIf CadIntermedia = "GET" Then
            .TipoPropiedad = etPropGet
          Else
            .TipoPropiedad = etPropSet
          End If
        Case "EVENT"
          .Tipo = etRutEvento
        Case "DECLARE"
          .Tipo = etRutDeclare
        'Palabras clave de rutinas DECLARE
        Case "LIB"
          CadIntermedia = guObjGeneral.QuitarParametro(CadTipoNombre, Chr$(34))
          .Libreria = guObjGeneral.QuitarParametro(CadTipoNombre, Chr$(34))
        Case "ALIAS"
          CadIntermedia = guObjGeneral.QuitarParametro(CadTipoNombre, Chr$(34))
          .Alias = guObjGeneral.QuitarParametro(CadTipoNombre, Chr$(34))
        'Nombre
        Case Else
          .Nombre = CadIntermedia
      End Select
    Wend
    'Obtiene la parte de parámetros
    .clParametros.Add CadParametros
    'Obtiene la parte de retorno
    If CadFinal <> "" Then
      If UCase$(Left$(CadFinal, 3)) = "AS " Then
        CadFinal = Right$(CadFinal, Len(CadFinal) - 3) 'Le quito el AS
      End If
      .TipoRetorno = guObjGeneral.QuitarParametro(CadFinal, "'")
      .Comentarios = CadFinal
    End If
  End With
  ReDim Preserve TablaRutinas(UBound(TablaRutinas) + 1)
End Sub

Public Sub addVariable(ByVal Cadena As String)
'--> Añade una variable a la subrutina (utiliza para ello la clase ClassVariables)
  TablaRutinas(UBound(TablaRutinas) - 1).clVariables.Add Cadena
End Sub

Public Sub addComentarios(ByVal Cadena As String)
'--> Añade un comentario a la rutina (utiliza para ello la clase ClassComentarios)
  TablaRutinas(UBound(TablaRutinas) - 1).clComentarios.Add Cadena
End Sub

Private Function getIndexHTML(ByRef clGeneradorHTML As ClassHTML, ByVal Tipo As etRutina) As String
'--> Obtiene una cadena HTML para el índice de las rutinas de tipo <B> Tipo </B>
Dim Indice As Integer
Dim ExistenValores As Boolean
Dim strCaption As String

  strCaption = IIf(Tipo = etRutFuncion, "Funciones", _
                   IIf(Tipo = etRutSubRutina, "Subrutinas", _
                                              IIf(Tipo = etRutEvento, "Eventos", _
                                                     IIf(Tipo = etRutDeclare, "Declare", _
                                                                              "Propiedades"))))
  getIndexHTML = getIndexHTML + "<TD CLASS=TableSubcaption> Nombre"
  getIndexHTML = getIndexHTML + "<TD CLASS=TableSubcaption> Ambito"
  If Tipo = etRutPropiedad Then
    getIndexHTML = getIndexHTML + "<TD CLASS=TableSubcaption> Tipo"
  End If
  If Tipo = etRutDeclare Then
    getIndexHTML = getIndexHTML + "<TD CLASS=TableSubcaption> Librería"
    getIndexHTML = getIndexHTML + "<TD CLASS=TableSubcaption> Alias"
  End If
  If Tipo <> etRutSubRutina And Tipo <> etRutEvento Then
    getIndexHTML = getIndexHTML + "<TD CLASS=TableSubcaption> Tipo Retorno"
  End If
  getIndexHTML = getIndexHTML + "<TD CLASS=TableSubcaption> Comentario"
  ExistenValores = False
  For Indice = 0 To UBound(TablaRutinas) - 1
    With TablaRutinas(Indice)
      If .Tipo = Tipo Then
        ExistenValores = True
        getIndexHTML = getIndexHTML + "<TR CLASS=TableRowColor>"
        If .Tipo <> etRutEvento Then
          getIndexHTML = getIndexHTML + "<TD CLASS=RowHeader>" + clGeneradorHTML.getHipertexto("#" + .Nombre, .Nombre)
        Else 'Los eventos no tienen cuerpo
          getIndexHTML = getIndexHTML + "<TD> " + .Nombre
        End If
        getIndexHTML = getIndexHTML + "<TD>" + getCadenaAmbito(.Ambito)
        If Tipo = etRutPropiedad Then
          getIndexHTML = getIndexHTML + "<TD>" + IIf(.TipoPropiedad = etPropLet, "Let", _
                                           IIf(.TipoPropiedad = etPropGet, "Get", "Set"))
        End If
        If Tipo = etRutDeclare Then
          getIndexHTML = getIndexHTML + "<TD>" + clGeneradorHTML.getCadenaVacia(.Libreria)
          getIndexHTML = getIndexHTML + "<TD>" + clGeneradorHTML.getCadenaVacia(.Alias)
        End If
        If Tipo <> etRutSubRutina And Tipo <> etRutEvento Then
          getIndexHTML = getIndexHTML + "<TD>" + clGeneradorHTML.getCadenaVacia(.TipoRetorno)
        End If
        getIndexHTML = getIndexHTML + "<TD>" + clGeneradorHTML.getCadenaVacia(.Comentarios)
      End If
    End With
  Next Indice
  getIndexHTML = IIf(ExistenValores, clGeneradorHTML.getTable(strCaption, getIndexHTML), "")
End Function

Private Sub writeHTMLRutinas(ByRef clGeneradorHTML As ClassHTML, ByVal blnConListaParametros As Boolean, _
                             ByVal blnConVariables As Boolean, ByVal NumeroFichero As Long)
'--> Escribe una cadena HTML para la descripción de las funciones
Dim Indice As Integer
Dim CadSalida As String

  Print #NumeroFichero, clGeneradorHTML.getAnchor("Rutinas")
  
  Print #NumeroFichero, "<TABLE WIDTH=100%> <TR WIDTH=100% BORDER=1 CLASS=Header>" + _
                        " <TD> Descripción de rutinas </TD> </TR> </TABLE>"
  
'  Print #NumeroFichero, "<H1 ALIGN = CENTER>" + _
'                        clGeneradorHTML.getCadenaConColor("FF0000", "Descripción de rutinas") + "</H1>"
                        
  For Indice = 0 To UBound(TablaRutinas) - 1
    With TablaRutinas(Indice)
      If .Tipo <> etRutEvento Then
        CadSalida = clGeneradorHTML.getAnchor(.Nombre)
        CadSalida = CadSalida + "<P CLASS=SubroutineHeader>" + .Nombre + "</P>"
        CadSalida = CadSalida + .CadenaCompleta + "&nbsp </P>"
        CadSalida = CadSalida + "<BLOCKQUOTE>"
        Print #NumeroFichero, CadSalida
        If blnConListaParametros Then
          .clParametros.writeHTML clGeneradorHTML, NumeroFichero
        End If
        If blnConVariables Then
          .clVariables.writeHTML clGeneradorHTML, NumeroFichero
        End If
        .clComentarios.writeHTML clGeneradorHTML, NumeroFichero
        Print #NumeroFichero, "</BLOCKQUOTE> <HR>"
      End If
    End With
  Next Indice
End Sub

Public Sub writeHTML(ByRef clGeneradorHTML As ClassHTML, ByVal blnConListaParametros As Boolean, ByVal blnConVariables As Boolean, ByVal NumeroFichero As Long)
'--> Escribe las rutinas del formulario sobre el fichero HTML
'--> @sub getIndexHTML
'--> @sub writeHTMLRutinas
Dim CadSalida As String

  CadSalida = getIndexHTML(clGeneradorHTML, etRutDeclare)
  CadSalida = CadSalida + getIndexHTML(clGeneradorHTML, etRutEvento)
  CadSalida = CadSalida + getIndexHTML(clGeneradorHTML, etRutSubRutina)
  CadSalida = CadSalida + getIndexHTML(clGeneradorHTML, etRutFuncion)
  CadSalida = CadSalida + getIndexHTML(clGeneradorHTML, etRutPropiedad)
  If CadSalida <> "" Then
    Print #NumeroFichero, CadSalida
    writeHTMLRutinas clGeneradorHTML, blnConListaParametros, blnConVariables, NumeroFichero
  End If
End Sub

Private Sub Class_Initialize()
  ReDim TablaRutinas(0)
End Sub

Private Sub Class_Terminate()
Dim Indice As Integer

  For Indice = 0 To UBound(TablaRutinas) - 1
    With TablaRutinas(Indice)
      Set .clParametros = Nothing
      Set .clVariables = Nothing
      Set .clComentarios = Nothing
    End With
  Next Indice
End Sub
