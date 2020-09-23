VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FR_Principal 
   Caption         =   "Documentador"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   Icon            =   "FR_Principal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin DocumentadorHTML.BarraEstado brEstado 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      Top             =   7770
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   476
   End
   Begin DocumentadorHTML.Spliter splVertical 
      Height          =   8445
      Left            =   720
      TabIndex        =   4
      Top             =   5610
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   14896
   End
   Begin ComCtl3.CoolBar clbHerramientas 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   953
      BandCount       =   2
      ImageList       =   "imlIcons"
      _CBWidth        =   8970
      _CBHeight       =   540
      _Version        =   "6.0.8169"
      Child1          =   "tlbHerramientas"
      MinHeight1      =   480
      Width1          =   3135
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "tlbExplorer"
      MinHeight2      =   450
      Width2          =   4365
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbExplorer 
         Height          =   450
         Left            =   3330
         TabIndex        =   3
         Top             =   45
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   794
         ButtonWidth     =   820
         ButtonHeight    =   794
         Appearance      =   1
         Style           =   1
         ImageList       =   "imlIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Back"
               Object.ToolTipText     =   "Atrás"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Forward"
               Object.ToolTipText     =   "Adelante"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Stop"
               Object.ToolTipText     =   "Detener"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Refresh"
               Object.ToolTipText     =   "Actualizar"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Home"
               Object.ToolTipText     =   "Inicio"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Search"
               Object.ToolTipText     =   "Búsqueda"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbHerramientas 
         Height          =   480
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   847
         ButtonWidth     =   820
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "imlIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Documentación"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Borrar"
               ImageIndex      =   9
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   720
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FR_Principal.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FR_Principal.frx":05EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FR_Principal.frx":08CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FR_Principal.frx":0BB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FR_Principal.frx":0E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FR_Principal.frx":1174
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FR_Principal.frx":1456
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FR_Principal.frx":199A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FR_Principal.frx":1EDE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lsoProyectos 
      Height          =   7080
      Left            =   60
      TabIndex        =   0
      Top             =   570
      Width           =   2445
   End
   Begin TabDlg.SSTab tabDocumentacion 
      Height          =   7065
      Left            =   2580
      TabIndex        =   5
      Top             =   600
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   12462
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Descripción"
      TabPicture(0)   =   "FR_Principal.frx":1FF6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblUltimaModificacion(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblUltimaModificacion(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblUltimaModificacion(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblUltimaModificacion(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "EDN_Archivo(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "EDN_Archivo(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "grdDocRelacionados"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkConVariables"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkConParametros"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ED_Nombre"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ED_Descripcion"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Documentación"
      TabPicture(1)   =   "FR_Principal.frx":2012
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "wbExplorer"
      Tab(1).Control(1)=   "shpWBExplorer"
      Tab(1).ControlCount=   2
      Begin VB.TextBox ED_Descripcion 
         Height          =   1155
         Left            =   1710
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Text            =   "FR_Principal.frx":202E
         Top             =   570
         Width           =   4485
      End
      Begin VB.TextBox ED_Nombre 
         Height          =   285
         Left            =   1710
         TabIndex        =   10
         Text            =   "Nombre del proyecto"
         Top             =   240
         Width           =   4485
      End
      Begin VB.CheckBox chkConParametros 
         Caption         =   "Con la definición de parámetros de las rutinas"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   180
         TabIndex        =   8
         Top             =   6030
         Value           =   1  'Checked
         Width           =   5775
      End
      Begin VB.CheckBox chkConVariables 
         Caption         =   "Con variables locales"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   6360
         Value           =   1  'Checked
         Width           =   5775
      End
      Begin MSFlexGridLib.MSFlexGrid grdDocRelacionados 
         Height          =   2835
         Left            =   150
         TabIndex        =   6
         Top             =   3120
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   5001
         _Version        =   393216
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin DocumentadorHTML.NombreArchivo EDN_Archivo 
         Height          =   345
         Index           =   0
         Left            =   270
         TabIndex        =   9
         Top             =   1740
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   609
         Caption         =   "Fichero de proyecto:"
         AutoSize        =   0   'False
         LabelSize       =   1455
      End
      Begin DocumentadorHTML.NombreArchivo EDN_Archivo 
         Height          =   375
         Index           =   1
         Left            =   270
         TabIndex        =   12
         Top             =   2400
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   661
         Caption         =   "Fichero Indice:"
         Cargar          =   0   'False
         AutoSize        =   0   'False
         LabelSize       =   1470
      End
      Begin SHDocVwCtl.WebBrowser wbExplorer 
         CausesValidation=   0   'False
         Height          =   4125
         Left            =   -74940
         TabIndex        =   19
         Top             =   90
         Width           =   4185
         ExtentX         =   7382
         ExtentY         =   7276
         ViewMode        =   1
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   -1  'True
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.Shape shpWBExplorer 
         Height          =   4245
         Left            =   -75000
         Top             =   0
         Width           =   4365
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción:"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   18
         Top             =   630
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   17
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label lblUltimaModificacion 
         Caption         =   "Ultima modificación "
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   16
         Top             =   2130
         Width           =   1815
      End
      Begin VB.Label lblUltimaModificacion 
         AutoSize        =   -1  'True
         Caption         =   "19-01-2000"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   5400
         TabIndex        =   15
         Top             =   2130
         Width           =   810
      End
      Begin VB.Label lblUltimaModificacion 
         Caption         =   "Ultima documentación"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   3600
         TabIndex        =   14
         Top             =   2820
         Width           =   1815
      End
      Begin VB.Label lblUltimaModificacion 
         AutoSize        =   -1  'True
         Caption         =   "19-01-2000"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   3
         Left            =   5430
         TabIndex        =   13
         Top             =   2820
         Width           =   810
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menú"
      Visible         =   0   'False
      Begin VB.Menu mnuMenuNuevoModificarBorrar 
         Caption         =   "&Nuevo"
         Index           =   0
      End
      Begin VB.Menu mnuMenuNuevoModificarBorrar 
         Caption         =   "&Modificar"
         Index           =   1
      End
      Begin VB.Menu mnuMenuNuevoModificarBorrar 
         Caption         =   "&Borrar"
         Index           =   2
      End
   End
End
Attribute VB_Name = "FR_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Ventana principal de la aplicación
Option Explicit

Private clProyecto As New ClassProyecto 'Clase principal del proyecto
Private clDocumentados As New ColDocumentados 'Clase con los proyectos documentados
'Private clStyle As New ClassEstilo
Private blnNuevo As Boolean 'Variable indicando si se está en un nuevo proyecto o con una modificación

Private Sub CargarDatos()
'--> Obtiene del ini los últimos ficheros leídos y estilos e inicializa la aplicación
  clDocumentados.readIni App.Path
  clDocumentados.loadList lsoProyectos
  'Inicializa la clase Estilo y las combos asociadas
'  clStyle.readIni App.Path
'  clStyle.loadComboNombres cboTag
'  clStyle.loadComboSize cboSize
End Sub

Private Sub GrabarDatos()
'--> Graba en el ini los últimos ficheros leídos
  clDocumentados.writeIni App.Path
'  clStyle.writeIni App.Path
End Sub

Private Sub initTag()
'--> Al cambiar alguno de los valores de color, tamaño, etc... debemos cambiar en la clase los valores almacenados
'--> Se utilizaba para cambiar el fichero de estilos, actualmente no se utiliza
'  clStyle.changeTag cboTag.ListIndex, edcEstilo.ColorFondo, edcEstilo.ColorTexto, _
'                    (chkTagBold.Value = vbChecked), (chkTagUnderline.Value = vbChecked), cboSize.Text
'  clStyle.seePreview wbExplorer(1)
End Sub

Private Sub Inhabilitar()
'--> Inhabilita / habilita la ventana para no permitir que se realicen acciones mientras se está aún trabajando
  FR_Principal.Enabled = Not FR_Principal.Enabled
End Sub

Private Sub AddProject(ByVal blnDesdeGrid As Boolean)
'--> Añade o modifica un proyecto documentado a la lista de proyectos
Dim intIndex As Integer

  If Not blnDesdeGrid Or (blnDesdeGrid And EDN_Archivo(0).NombreArchivo <> "" And EDN_Archivo(1).NombreArchivo <> "") Then
    If blnNuevo Then
      intIndex = lsoProyectos.ListCount + 1
    Else
      intIndex = lsoProyectos.ListIndex
    End If
    clDocumentados.Modify intIndex, (chkConVariables.Value = vbChecked), (chkConParametros.Value = vbChecked), _
                          Format(Now, ConstCadenaFecha), EDN_Archivo(1).NombreArchivo, _
                          EDN_Archivo(0).NombreArchivo, ED_Descripcion.Text, ED_Nombre.Text, grdDocRelacionados
  End If
End Sub

Private Sub borrarFicheros()
'--> Borra los ficheros del directorio donde se va a generar la documentación (no es recursivo)
  On Error GoTo ErrorBorrado
    Kill EDN_Archivo(1).Path & "*.*"
  Exit Sub
  
ErrorBorrado:
End Sub

Private Sub GenerarDocumentacion()
'--> Botón para comenzar la documentación.
'--> Utiliza la clase ClassProyecto para realizar todo el proceso, al terminar cambia la URL del navegador para ver la documentación
'--> @sub Inhabilitar
'--> @sub NavegarA
Dim Indice As Integer, NumeroFichero As Integer

  If Trim(UCase$(EDN_Archivo(0).NombreArchivo)) = Trim(UCase$(EDN_Archivo(1).NombreArchivo)) Then
    guObjGeneral.MensajeError "Los nombres de los ficheros deben ser diferentes.", mErrError
  Else
    tabDocumentacion.Tab = 1
    Screen.MousePointer = vbHourglass
    If guObjGeneral.MensajeError("¿Desea borrar los ficheros del directorio '" + EDN_Archivo(1).Path + "'?", mErrInterrogacion) Then
      borrarFicheros
    End If
    Inhabilitar
'    clStyle.writeStyleCSS EDN_Archivo(1).Path
    With clProyecto
      .Nombre = ED_Nombre
      .Descripcion = ED_Descripcion
      .DirectorioDocumentacion = EDN_Archivo(1).Path
      .ConParametros = (chkConParametros.Value = vbChecked)
      .ConVariables = (chkConVariables.Value = vbChecked)
      .LeerProyecto EDN_Archivo(0).NombreArchivo
      .writeHTML EDN_Archivo(1).NombreArchivo, clDocumentados.Item(lsoProyectos.ListIndex + 1)
    End With
    NavegarA clProyecto.DirectorioDocumentacion + "main_doc.html"
    Inhabilitar
    Screen.MousePointer = vbDefault
  End If
End Sub

Private Sub NavegarA(ByVal strURL As String)
'--> Posiciona el navegador en una URL de fichero
  On Error Resume Next
  With wbExplorer
    .FullScreen = True
    .Navigate "file://" + strURL
    .Refresh
    .TheaterMode = False
  End With
End Sub

Private Sub cboSize_Change()
'--> Para cambiar los estilos del fichero CSS, actualmente no se utiliza
  initTag
End Sub

Private Sub cboTag_Click()
'--> Al cambiar en la combo el tag debemos cambiar los valores, se utilizaba para cambiar el CSS actualmente no se utiliza
Dim colFondo As OLE_COLOR, colTexto As OLE_COLOR
Dim blnBold As Boolean, blnUnderline As Boolean
Dim strSize As String

'  clStyle.getTag cboTag.ListIndex, colFondo, colTexto, blnBold, blnUnderline, strSize
'  cboSize.Text = strSize
'  chkTagBold.Value = IIf(blnBold, vbChecked, vbUnchecked)
'  chkTagUnderline.Value = IIf(blnUnderline, vbChecked, vbUnchecked)
'  edcEstilo.ColorFondo = colFondo
'  edcEstilo.ColorTexto = colTexto
End Sub

Private Sub clbHerramientas_HeightChanged(ByVal NewHeight As Single)
  splVertical_Redimensionar splVertical.Posicion
End Sub

Private Sub chkTagBold_Click()
'--> Para cambiar los estilos del fichero CSS, actualmente no se utiliza
  initTag
End Sub

Private Sub chkTagUnderline_Click()
'--> Para cambiar los estilos del fichero CSS, actualmente no se utiliza
  initTag
End Sub

Private Sub edcEstilo_Change()
'--> Para cambiar los estilos del fichero CSS, actualmente no se utiliza
  initTag
End Sub

Private Sub Form_Load()
  EDN_Archivo(0).Filter = "Archivo de Proyectos (*.vbp) |*.vbp| Todos los Archivos (*.*) |*.*"
  EDN_Archivo(1).Filter = "Archivo HTML (*.html) |*.html| Todos los Archivos (*.*) |*.*"
  EDN_Archivo(0).LabelSize = EDN_Archivo(1).LabelSize
  With grdDocRelacionados
    .Rows = 2
    .Cols = 2
    .TextMatrix(0, 0) = "Descripción"
    .TextMatrix(0, 1) = "Fichero"
  End With
  With splVertical
    .Top = 0
    .Left = 0
    .Redimensionar Me.Width, Me.Height
    .ZOrder 1
  End With
  guObjGeneral.Inicializar "Documentador de VB", Nothing
  CargarDatos
  wbExplorer.GoHome
  blnNuevo = False
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    If Me.Width < 9090 Then
      Me.Width = 9090
    ElseIf Me.Height < 6465 Then
      Me.Height = 6465
    Else
      splVertical_Redimensionar splVertical.Posicion
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  GrabarDatos
  Set guObjGeneral = Nothing
  Set clProyecto = Nothing
  Set clDocumentados = Nothing
End Sub

Private Sub grdDocRelacionados_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'--> Al pulsar con el botón derecho sobre el grid muestra el menú de modificaciones
  If Button = vbRightButton And grdDocRelacionados.Row <> 0 Then
    PopupMenu mnuMenu
  End If
End Sub

Private Sub lsoProyectos_Click()
'--> Al posicionarse sobre otro proyecto debe cargar los datos de éste (nombre, parámetros... y navegar hacia su documentación)
Dim objDocumentado As ClassDocumentado

  If lsoProyectos.ListCount > 0 Then
    Set objDocumentado = clDocumentados.Item(lsoProyectos.ListIndex + 1)
    With objDocumentado
      ED_Nombre.Text = .strName
      ED_Descripcion.Text = .strDescription
      EDN_Archivo(0).NombreArchivo = .strFileProject
      EDN_Archivo(1).NombreArchivo = .strFileIndex
      lblUltimaModificacion(3).Caption = .strDateGeneration
      chkConParametros.Value = IIf(.blnWithParameters, vbChecked, vbUnchecked)
      chkConVariables.Value = IIf(.blnWithVariables, vbChecked, vbUnchecked)
      NavegarA EDN_Archivo(1).Path + "main_doc.html"
    End With
    Set objDocumentado = Nothing
    clDocumentados.loadGrid lsoProyectos.ListIndex + 1, grdDocRelacionados
    blnNuevo = False
  End If
End Sub

Private Sub mnuMenuNuevoModificarBorrar_Click(Index As Integer)
'--> Opción de nuevo, modificar o borrar un documento relacionado
'--> Pasa en gstrDatoIntermedio el comentario del fichero y su dirección separados por @ y espera la devolución de valores de _
     la misma forma
'--> En gAccion pasa nuevo o modificar, los borrados se hacen directametne desde esta ventana
'--> @form frmDocumentosRelacionados
Dim intRow As Integer

  With grdDocRelacionados
    intRow = .Row
    Select Case Index
      Case 0 'Nuevo
        guAccion = etAcNuevo
        guObjGeneral.AbrirVentanaModal frmDocumentosRelacionados
        If guAccion = etAcRealizado Then
          .addItem guObjGeneral.QuitarParametro(gstrDatoIntermedio, "@") & vbTab & gstrDatoIntermedio
          If .Rows > 2 Then
            If .TextMatrix(1, 0) = "" Then .RemoveItem 1
          End If
        End If
      Case 1 'Modificar
        If .TextMatrix(intRow, 0) <> "" Then
          guAccion = etAcModificar
          gstrDatoIntermedio = .TextMatrix(intRow, 0) & "@" & .TextMatrix(intRow, 1)
          guObjGeneral.AbrirVentanaModal frmDocumentosRelacionados
          If guAccion = etAcRealizado Then
            .TextMatrix(intRow, 0) = guObjGeneral.QuitarParametro(gstrDatoIntermedio, "@")
            .TextMatrix(intRow, 1) = gstrDatoIntermedio
          End If
        End If
      Case 2 'Borrar
        If .TextMatrix(intRow, 0) <> "" Then
          If guObjGeneral.MensajeError("¿Realmente desea borrar el documento '" + .TextMatrix(intRow, 0) + "'?", mErrInterrogacion) Then
            If .Rows > 2 Then
              .RemoveItem intRow
            Else
              .TextMatrix(1, 0) = ""
              .TextMatrix(1, 1) = ""
            End If
          End If
        End If
    End Select
  End With
  If guAccion = etAcRealizado Then AddProject True
End Sub

Private Sub splVertical_Redimensionar(ByVal x As Integer)
Dim intTabAnterior As Integer

  On Error Resume Next
  With lsoProyectos
    .Top = FR_Principal.ScaleTop + clbHerramientas.Height + 20
    .Width = x - .Left
    .Height = FR_Principal.ScaleHeight - .Top - brEstado.Height
  End With
  intTabAnterior = tabDocumentacion.Tab
  With tabDocumentacion
    .Top = lsoProyectos.Top
    .Left = x + 70
    .Height = lsoProyectos.Height
    .Width = FR_Principal.ScaleWidth - .Left
  End With
  tabDocumentacion.Tab = 0
  ED_Nombre.Width = tabDocumentacion.Width - ED_Nombre.Left - 150
  ED_Descripcion.Width = ED_Nombre.Width
  EDN_Archivo(0).Width = tabDocumentacion.Width - EDN_Archivo(0).Left - 150
  EDN_Archivo(1).Width = EDN_Archivo(0).Width
  With grdDocRelacionados
    .Width = EDN_Archivo(0).Width
    .ColWidth(0) = .Width / 4
    .ColWidth(1) = .Width - 3 * .Width / 4
  End With
  chkConParametros.Width = EDN_Archivo(0).Width
  chkConVariables.Width = EDN_Archivo(0).Width
  lblUltimaModificacion(1).Left = tabDocumentacion.Width - lblUltimaModificacion(1).Width - 150
  lblUltimaModificacion(0).Left = lblUltimaModificacion(1).Left - lblUltimaModificacion(0).Width
  lblUltimaModificacion(3).Left = lblUltimaModificacion(1).Left
  lblUltimaModificacion(2).Left = lblUltimaModificacion(3).Left - lblUltimaModificacion(2).Width
  tabDocumentacion.Tab = 1
  'Diseño de la hoja de estilos
  tabDocumentacion.Tab = 2
  With wbExplorer
    .Height = tabDocumentacion.Height - tabDocumentacion.TabHeight - .Top - 100
    .Width = tabDocumentacion.Width - .Left - 100
  End With
  With shpWBExplorer
    .Top = wbExplorer.Top - 2
    .Left = wbExplorer.Left - 2
    .Width = wbExplorer.Width + 2
    .Height = wbExplorer.Height + 2
    .ZOrder 0
  End With
  tabDocumentacion.Tab = intTabAnterior
End Sub

Private Sub tabDocumentacion_Click(PreviousTab As Integer)
  If tabDocumentacion.Tab = 2 Then
    If EDN_Archivo(1).NombreArchivo <> "" Then
      NavegarA EDN_Archivo(1).Path + "main_doc.html"
    End If
  End If
End Sub

Private Sub tlbExplorer_ButtonClick(ByVal Button As MSComctlLib.Button)
'--> Acciones de la barra de herramientas del Navegador
  On Error Resume Next
  Select Case Button.Key
    Case "Back"
      wbExplorer(0).GoBack
    Case "Forward"
      wbExplorer(0).GoForward
    Case "Refresh"
      wbExplorer(0).Refresh
    Case "Home"
      wbExplorer(0).GoHome
    Case "Search"
      wbExplorer(0).GoSearch
    Case "Stop"
      wbExplorer(0).Stop
  End Select
End Sub

Private Sub tlbHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
'--> Acciones de la barra de herramientas del documentador (nuevo proyecto, generar documentación, eliminar proyecto)
  Select Case Button.Index
    Case 1 'Nuevo
      blnNuevo = True
      ED_Nombre.Text = ""
      ED_Descripcion.Text = ""
    Case 2 'Generar documentación
      AddProject False
      GenerarDocumentacion
      clDocumentados.loadList lsoProyectos
      blnNuevo = False
    Case 3 'Borrar
      If lsoProyectos.ListCount > 0 Then
        guObjGeneral.MensajeError "No se borrarán los ficheros de documentación", mErrInformacion
        clDocumentados.Remove lsoProyectos.ListIndex + 1
        lsoProyectos.RemoveItem lsoProyectos.ListIndex
      End If
  End Select
End Sub
