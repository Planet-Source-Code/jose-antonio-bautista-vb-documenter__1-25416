VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl EdColor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3225
   PropertyPages   =   "CTL_ColorFondoTexto.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   3225
   Begin MSComDlg.CommonDialog DLG_Color 
      Left            =   2088
      Top             =   120
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton BT_CambioColor 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   1
      Left            =   1416
      TabIndex        =   2
      Top             =   24
      Width           =   372
   End
   Begin VB.CommandButton BT_CambioColor 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   0
      Left            =   24
      TabIndex        =   1
      Top             =   24
      Width           =   324
   End
   Begin VB.Label LB_Prueba 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Prueba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   48
      TabIndex        =   0
      Top             =   24
      Width           =   1740
   End
End
Attribute VB_Name = "EdColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'--> Control de usuario para cambiar los colores y fuentes
Option Explicit

Public Event Change()
Attribute Change.VB_Description = "Evento producido al cambiar el usuario el color de fondo y/o texto"

Public Property Let Caption(ByVal pCaption As String)
  LB_Prueba = pCaption
  PropertyChanged
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Texto"
Attribute Caption.VB_ProcData.VB_Invoke_Property = "StandardColor;Apariencia"
Attribute Caption.VB_MemberFlags = "200"
  Caption = LB_Prueba
End Property

Public Property Let ColorFondo(ByVal pColor As OLE_COLOR)
Attribute ColorFondo.VB_Description = "Color de fondo"
Attribute ColorFondo.VB_ProcData.VB_Invoke_PropertyPut = "StandardColor;Apariencia"
  LB_Prueba.BackColor = pColor
  PropertyChanged
End Property

Public Property Get ColorFondo() As OLE_COLOR
  ColorFondo = LB_Prueba.BackColor
End Property

Public Property Let ColorTexto(ByVal pColor As OLE_COLOR)
Attribute ColorTexto.VB_Description = "Color de texto"
Attribute ColorTexto.VB_ProcData.VB_Invoke_PropertyPut = "StandardColor;Apariencia"
  LB_Prueba.ForeColor = pColor
  PropertyChanged
End Property

Public Property Get ColorTexto() As OLE_COLOR
  ColorTexto = LB_Prueba.ForeColor
End Property

Private Sub BT_CambioColor_Click(Index As Integer)
  On Error GoTo Cancelado
  With DLG_Color
    .Flags = cdlCCPreventFullOpen Or cdlCCRGBInit
    .CancelError = True
    .COLOR = IIf(Index = 0, LB_Prueba.BackColor, LB_Prueba.ForeColor)
    .ShowColor
    If Index = 0 Then
      ColorFondo = .COLOR
    Else
      ColorTexto = .COLOR
    End If
  End With
  RaiseEvent Change
  
Cancelado:
End Sub

Private Sub UserControl_InitProperties()
  ColorFondo = vbButtonFace
  ColorTexto = vbButtonText
  Caption = "Prueba"
End Sub

Private Sub UserControl_Resize()
  If Height <> 252 Then
    Height = 252
  ElseIf Width < 1356 Then
    Width = 1356
  Else
    With BT_CambioColor(0)
      .Top = 0
      .Width = 276
      .Left = 0
      .Height = Height
    End With
    With BT_CambioColor(1)
      .Top = 0
      .Width = 276
      .Left = Width - .Width
      .Height = Height
    End With
    With LB_Prueba
      .Top = 0
      .Left = 0
      .Width = Width
      .Height = Height
    End With
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ColorFondo = PropBag.ReadProperty("ColorFondo", vbButtonFace)
  ColorTexto = PropBag.ReadProperty("ColorTexto", vbButtonText)
  Caption = PropBag.ReadProperty("Caption", "Prueba")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "ColorFondo", ColorFondo, vbButtonFace
  PropBag.WriteProperty "ColorTexto", ColorTexto, vbButtonText
  PropBag.WriteProperty "Caption", Caption, "Prueba"
End Sub
