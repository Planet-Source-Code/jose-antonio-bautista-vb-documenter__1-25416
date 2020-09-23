VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl NombreArchivo 
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   LockControls    =   -1  'True
   ScaleHeight     =   465
   ScaleWidth      =   6495
   ToolboxBitmap   =   "CTL_NombreArchivo.ctx":0000
   Begin MSComDlg.CommonDialog DLG_Ficheros 
      Left            =   5784
      Top             =   24
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton BT_ObtenerFichero 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5016
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   48
      Width           =   396
   End
   Begin VB.TextBox ED_Archivo 
      Height          =   300
      Left            =   1392
      TabIndex        =   1
      Top             =   48
      Width           =   3540
   End
   Begin VB.Label LB_Titulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Hoja:"
      ForeColor       =   &H8000000D&
      Height          =   192
      Left            =   120
      TabIndex        =   0
      Top             =   72
      Width           =   1008
   End
End
Attribute VB_Name = "NombreArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> Control para recibir un nombre de fichero a grabar o salvar
Option Explicit

Private FuncionCargar As Boolean 'Indice si se desea abrir el diálogo de directorios para leer o escribir
Private Filtro As String 'Tener en cuenta que el DLG_Ficheros lo pueden compartir varios controles con filtros diferentes

Public Function Path() As String
'--> Consigue el directorio a partir de un nombre de fichero con todo el path.
Dim ObjFichero As New ClassFicheros

  Path = ObjFichero.ObtenerPath(ED_Archivo)
  Set ObjFichero = Nothing
End Function

Public Property Let Caption(ByVal Titulo As String)
Attribute Caption.VB_Description = "Título del control"
  LB_Titulo = Titulo
  UserControl_Resize
  PropertyChanged
End Property

Public Property Get Caption() As String
  Caption = LB_Titulo
End Property

Public Property Let Cargar(ByVal VarCargar As Boolean)
Attribute Cargar.VB_Description = "Determina si el control se utiliza para Cargar (True) o Salvar (False) archivos"
  FuncionCargar = VarCargar
  PropertyChanged
End Property

Public Property Get Cargar() As Boolean
  Cargar = FuncionCargar
End Property

Public Property Let NombreArchivo(ByVal Nombre As String)
Attribute NombreArchivo.VB_Description = "Nombre del archivo por defecto"
  ED_Archivo = Nombre
  PropertyChanged
End Property

Public Property Get NombreArchivo() As String
  NombreArchivo = Trim(ED_Archivo)
End Property

Public Property Let AutoSize(ByVal Auto As Boolean)
Attribute AutoSize.VB_Description = "Determina si el nombre del control se redimensiona al cambiar su contenido"
  LB_Titulo.AutoSize = Auto
  PropertyChanged
End Property

Public Property Get AutoSize() As Boolean
  AutoSize = LB_Titulo.AutoSize
End Property

Public Property Let LabelSize(ByVal Ancho As Integer)
Attribute LabelSize.VB_Description = "Tamaño del título"
  LB_Titulo.AutoSize = False
  LB_Titulo.Width = Ancho
  UserControl_Resize
  PropertyChanged
End Property

Public Property Get LabelSize() As Integer
  LabelSize = LB_Titulo.Width
End Property

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal pEnabled As Boolean)
  UserControl.Enabled = pEnabled
  PropertyChanged "ENABLED"
End Property

Public Property Let Filter(ByVal pFiltro As String)
Attribute Filter.VB_Description = "Filtro que utilizará al cargar / salvar archivos"
  Filtro = pFiltro
  PropertyChanged
End Property

Public Property Get Filter() As String
  Filter = Filtro
End Property

Public Sub WHelpId(ByVal pFicheroAyuda As String, ByVal pIdAyuda As Integer)
Attribute WHelpId.VB_Description = "Id de ayuda del control"
'--> Cambia el WhatsThisHelpId del control
  App.HelpFile = pFicheroAyuda
  BT_ObtenerFichero.WhatsThisHelpID = pIdAyuda
  ED_Archivo.WhatsThisHelpID = pIdAyuda
End Sub

Private Sub BT_ObtenerFichero_Click()
'--> Botón para conseguir los nombres de los archivos
Dim ObjFichero As New ClassFicheros

  ED_Archivo = ObjFichero.DLGNombreFichero(DLG_Ficheros, Cargar, ED_Archivo, Filtro)
  Set ObjFichero = Nothing
End Sub

Private Sub UserControl_InitProperties()
  Caption = "Nombre Hoja:"
  Cargar = True
  NombreArchivo = ""
  AutoSize = True
  LabelSize = 300
  Enabled = True
  Filter = ""
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Caption = PropBag.ReadProperty("Caption", "Nombre Hoja:")
  Cargar = PropBag.ReadProperty("Cargar", True)
  NombreArchivo = PropBag.ReadProperty("NombreArchivo", "")
  AutoSize = PropBag.ReadProperty("AutoSize", True)
  LabelSize = PropBag.ReadProperty("LabelSize", 300)
  Enabled = PropBag.ReadProperty("Enabled", True)
  Filter = PropBag.ReadProperty("Filter", "")
End Sub

Private Sub UserControl_Resize()
  LB_Titulo.Left = 10
  BT_ObtenerFichero.Left = Width - BT_ObtenerFichero.Width - 30
  ED_Archivo.Left = LB_Titulo.Left + LB_Titulo.Width + 40
  ED_Archivo.Width = Width - LB_Titulo.Width - BT_ObtenerFichero.Width - 150
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Caption", Caption, "Nombre Hoja:"
  PropBag.WriteProperty "Cargar", Cargar, True
  PropBag.WriteProperty "NombreArchivo", NombreArchivo, ""
  PropBag.WriteProperty "AutoSize", AutoSize, True
  PropBag.WriteProperty "LabelSize", LabelSize, 300
  PropBag.WriteProperty "Enabled", Enabled, True
  PropBag.WriteProperty "Filter", Filtro, ""
End Sub
