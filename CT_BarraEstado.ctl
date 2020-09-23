VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl BarraEstado 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
   Enabled         =   0   'False
   FillColor       =   &H00808080&
   MaskColor       =   &H80000010&
   ScaleHeight     =   270
   ScaleWidth      =   9375
   ToolboxBitmap   =   "CT_BarraEstado.ctx":0000
   Begin DocumentadorHTML.BarraProgreso brProgreso 
      Height          =   195
      Left            =   1290
      Top             =   60
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   344
      Max             =   0
   End
   Begin MSComctlLib.StatusBar ST_Barra 
      Height          =   204
      Left            =   6528
      TabIndex        =   1
      Top             =   24
      Width           =   2796
      _ExtentX        =   4948
      _ExtentY        =   344
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   714
            MinWidth        =   706
            Text            =   "May"
            TextSave        =   "May"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   714
            MinWidth        =   706
            Text            =   "Num"
            TextSave        =   "Num"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   714
            MinWidth        =   706
            Text            =   "Ins"
            TextSave        =   "Ins"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   714
            MinWidth        =   706
            Text            =   "Scr"
            TextSave        =   "Scr"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   706
            TextSave        =   "15:12"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "16/03/00"
         EndProperty
      EndProperty
   End
   Begin VB.Label LB_Mensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preparado ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   192
      Left            =   48
      TabIndex        =   0
      Top             =   24
      Width           =   1092
   End
End
Attribute VB_Name = "BarraEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> Control de la barra de estado de la ventana principal
Option Explicit

Private Const MensajeDefecto = "Preparado ... "

Private CadenaPreparado As String 'Cadena que muestra en la barra de estado

Property Let CambiarHelpId(ByVal InterIdAyuda As Integer)
Attribute CambiarHelpId.VB_Description = "HID del control."
  LB_Mensaje.WhatsThisHelpID = InterIdAyuda
  brProgreso.WhatsThisHelpID = InterIdAyuda
  ST_Barra.WhatsThisHelpID = InterIdAyuda
  PropertyChanged
End Property

Property Get CambiarHelpId() As Integer
  CambiarHelpId = LB_Mensaje.WhatsThisHelpID
End Property

Private Sub CambiarPosiciones()
'--> Cambia los tamaños de la barra de progreso y la etiqueta cuando está o no visible
  On Error Resume Next 'Pueden darse errores si la ventana es demasiado pequeña
  If brProgreso.Visible Then
    brProgreso.Width = Width - ST_Barra.Width - LB_Mensaje.Width - 300
    brProgreso.Left = LB_Mensaje.Width + LB_Mensaje.Left + 25
  End If
  ST_Barra.Refresh
  LB_Mensaje.Refresh
End Sub

Public Sub EscribirMensajeBarraEstado(Optional ByVal Cadena As String = "")
'--> Escribe un mensaje sobre la barra de estado
  LB_Mensaje = IIf(Cadena <> "", Cadena + " ", CadenaPreparado)
End Sub

Public Sub BarraProgresoCambiar()
'--> Incrementa en uno la barra de progreso
  brProgreso.Cambiar
End Sub

Public Sub BarraProgresoCerrar()
'--> Cierra la barra de progreso, deja como mensaje el establecido en <B> CadenaPreparado </B>
  brProgreso.Visible = False
  CambiarPosiciones
  EscribirMensajeBarraEstado
End Sub

Public Sub BarraProgresoInicializar(ByVal Cadena As String, ByVal Minimo As Long, ByVal Maximo As Long)
'--> Inicializa la barra de progreso.
  On Error Resume Next
  EscribirMensajeBarraEstado Cadena
  brProgreso.Inicializar Minimo, Maximo
  brProgreso.Visible = True
  CambiarPosiciones
End Sub

Property Let Mensaje(ByVal Cadena As String)
Attribute Mensaje.VB_Description = "Mensaje por defecto de la barra de estado."
'--> Define un mensaje como mensaje por defecto sobre la barra de estado.
  CadenaPreparado = Cadena
  EscribirMensajeBarraEstado CadenaPreparado
  PropertyChanged
End Property

Property Get Mensaje() As String
'--> Obtiene un mensaje por defecto
  Mensaje = CadenaPreparado
End Property

Private Sub UserControl_Initialize()
  EscribirMensajeBarraEstado CadenaPreparado
  brProgreso.Visible = False
End Sub

Private Sub UserControl_InitProperties()
  CambiarHelpId = 0
  Mensaje = MensajeDefecto
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  CambiarHelpId = PropBag.ReadProperty("CambiarHelpId", 0)
  Mensaje = PropBag.ReadProperty("Mensaje", MensajeDefecto)
End Sub

Private Sub UserControl_Resize()
  If Height < 264 Then
    Height = 264
  Else
    LB_Mensaje.Top = (Height - LB_Mensaje.Height) / 2 - 10
    brProgreso.Top = (Height - brProgreso.Height) / 2 - 25
    With ST_Barra
      .Top = (Height - .Height) / 2 - 25
      .Left = Width - .Width - 70
    End With
    CambiarPosiciones
  End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "CambiarHelpId", LB_Mensaje.WhatsThisHelpID, 0
  PropBag.WriteProperty "Mensaje", CadenaPreparado, MensajeDefecto
End Sub
