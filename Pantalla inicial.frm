VERSION 5.00
Begin VB.Form FR_PantallaInicial 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3135
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Pantalla inicial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   543
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2988
      Left            =   48
      TabIndex        =   0
      Top             =   24
      Width           =   8010
      Begin VB.Timer TM_Temporizador 
         Interval        =   5000
         Left            =   3576
         Top             =   864
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         Caption         =   "Versión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   3384
         TabIndex        =   1
         Top             =   2088
         Width           =   4524
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1800
         Left            =   3360
         TabIndex        =   3
         Top             =   216
         Width           =   4548
      End
      Begin VB.Label lblCopyright 
         Alignment       =   2  'Center
         Caption         =   "@ G.I.S.S. 1.998"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   3360
         TabIndex        =   2
         Top             =   2616
         Width           =   4572
      End
      Begin VB.Image ImgLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   2652
         Left            =   120
         Stretch         =   -1  'True
         Top             =   192
         Width           =   3192
      End
   End
End
Attribute VB_Name = "FR_PantallaInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Pantalla inicial de la aplicación
Option Explicit

Private Sub Form_Click()
'--> Todos los Click llaman a ImgLogo_Click que directamente descarga el formulario
  ImgLogo_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  ImgLogo_Click
End Sub

Private Sub Form_Load()
  'ImgLogo.Picture = LoadResPicture(ConstIconoLogoGiss, vbResBitmap)
End Sub

Private Sub Frame1_Click()
  ImgLogo_Click
End Sub

Private Sub ImgLogo_Click()
  Unload Me
End Sub

Private Sub lblCopyright_Click()
  ImgLogo_Click
End Sub

Private Sub lblProductName_Click()
  ImgLogo_Click
End Sub

Private Sub lblVersion_Click()
  ImgLogo_Click
End Sub

Private Sub TM_Temporizador_Timer()
  ImgLogo_Click
End Sub
