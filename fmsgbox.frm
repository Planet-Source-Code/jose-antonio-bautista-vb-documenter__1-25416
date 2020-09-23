VERSION 5.00
Begin VB.Form FR_MsgBox 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administración de Metadatos"
   ClientHeight    =   2280
   ClientLeft      =   2070
   ClientTop       =   2370
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "fmsgbox.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2280
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1668
      Left            =   96
      TabIndex        =   2
      Top             =   24
      Width           =   5724
      Begin VB.Label LB_Mensaje 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "LLLL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1308
         Left            =   936
         TabIndex        =   3
         Top             =   240
         Width           =   4620
      End
      Begin VB.Image IM_Error 
         Appearance      =   0  'Flat
         Height          =   600
         Left            =   144
         Top             =   240
         Width           =   672
      End
   End
   Begin VB.CommandButton BT_AceptarCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   3288
      TabIndex        =   1
      Top             =   1812
      Width           =   1788
   End
   Begin VB.CommandButton BT_AceptarCancelar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   1128
      TabIndex        =   0
      Top             =   1824
      Width           =   1788
   End
End
Attribute VB_Name = "FR_MsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Ventana para la presentación de mensajes de error.
'--> Dependiendo de gstrDatoIntermedio muestra un tipo de mensaje u otro.
'--> Valores de entrada en gstrDatoIntermedio
'--> <BLOCKQUOTE>
'--> 0: Mensaje de error informativo.
'--> 1: Mensaje de error Aceptar / Cancelar. Devuelve en <B> gstrDatoIntermedio </B>:
'--> <BLOCKQUOTE>
'--> False: Se pulsó el botón de cancelar.
'--> True: Se pulsó el botón de aceptar.
'--> </BLOCKQUOTE>
'--> 2: Mensaje de error exclamativo.
'--> 3: Mensaje de error corte comunicaciones
'--> 4: Mensaje de error de base de datos
'--> </BLOCKQUOTE>
'--> Normalmente para presentar esta ventana se llama a la rutina <B> MensajeError </B> de la clase <B> ClassRutinasGenerales </B>
Option Explicit

Private Const SeparacionBotones As Integer = 500 'Separación de los botones de aceptar y cancelar

Private Sub BT_AceptarCancelar_Click(Index As Integer)
  gstrDatoIntermedio = Str$((Index = 0))
  Unload Me
End Sub

Private Sub Form_Load()
'--> Se encarga de según el dato que venga en gstrDatoIntermedio mostrar un icono u otro _
     y cambiar los botones visibles o invisibles (sólo mErrInterrogacion tiene botón Cancelar)
  Select Case Val(gstrDatoIntermedio)
    Case mErrInformacion, mErrError, mErrComunicacion, mErrErrorBaseDatos
      Select Case Val(gstrDatoIntermedio)
        Case mErrInformacion
          IM_Error.Picture = LoadResPicture(ConstIconoInformacion, vbResIcon)
        Case mErrError
          IM_Error.Picture = LoadResPicture(ConstIconoExclamacion, vbResIcon)
        Case mErrComunicacion
          IM_Error.Picture = LoadResPicture(ConstIconoCorteComunicacion, vbResIcon)
        Case mErrErrorBaseDatos
          IM_Error.Picture = LoadResPicture(ConstIconoBaseDatos, vbResIcon)
      End Select
      With BT_AceptarCancelar(0)
        .Left = (FR_MsgBox.ScaleWidth - .Width) / 2
        .Default = True
        .Cancel = True
      End With
      BT_AceptarCancelar(1).Visible = False
    Case mErrInterrogacion
      IM_Error.Picture = LoadResPicture(ConstIconoInterrogacion, vbResIcon)
      With BT_AceptarCancelar(0)
        .Left = (FR_MsgBox.ScaleWidth - .Width - BT_AceptarCancelar(1).Width - SeparacionBotones) / 2
        .Default = True
      End With
      With BT_AceptarCancelar(1)
        .Left = BT_AceptarCancelar(0).Left + BT_AceptarCancelar(0).Width + SeparacionBotones
        .Visible = True
        .Cancel = True
      End With
  End Select
  gstrDatoIntermedio = Str$(False)
  Screen.MousePointer = vbDefault
End Sub
