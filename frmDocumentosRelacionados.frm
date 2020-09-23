VERSION 5.00
Begin VB.Form frmDocumentosRelacionados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documento relacionado"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   1785
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdAceptarCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Index           =   1
      Left            =   3248
      TabIndex        =   1
      Top             =   990
      Width           =   1485
   End
   Begin VB.CommandButton cmdAceptarCancelar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   1530
      TabIndex        =   0
      Top             =   990
      Width           =   1485
   End
   Begin DocumentadorHTML.NombreArchivo EDN_Archivo 
      Height          =   345
      Left            =   248
      TabIndex        =   3
      Top             =   480
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   609
      Caption         =   "&Nombre de Fichero:"
      AutoSize        =   0   'False
      LabelSize       =   1470
   End
   Begin VB.Label Label1 
      Caption         =   "&Descripción:"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   248
      TabIndex        =   2
      Top             =   180
      Width           =   1845
   End
End
Attribute VB_Name = "frmDocumentosRelacionados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Ventana para introducir los datos de ficheros relacionados con la documentacion
'--> En guAccion se le pasa la acción a realizar, devuelve Realizado o Cancelado
'--> En gstrDatoIntermedio se le pasa una cadena formada por la descripción, un carácter '@' y el nombre de fichero que _
     la ventana devuelve con los cambios con el mismo formato
Option Explicit

Private Sub cmdAceptarCancelar_Click(Index As Integer)
'--> Si el Index = 0 (aceptado) se comprueba que se han introducido los datos y se pasan en gstrDatoIntermedio
  If Index = 0 Then
    If Trim(txtDescripcion.Text) = "" Then
      guObjGeneral.MensajeError "Introduzca la descripción del fichero", mErrError
    ElseIf EDN_Archivo.NombreArchivo = "" Then
      guObjGeneral.MensajeError "Introduzca el nombre del archivo", mErrError
    Else
      gstrDatoIntermedio = txtDescripcion.Text + "@" + EDN_Archivo.NombreArchivo
      guAccion = etAcRealizado
    End If
  Else
    guAccion = etAcCancelado
  End If
  If guAccion = etAcCancelado Or guAccion = etAcRealizado Then Unload Me
End Sub

Private Sub Form_Load()
  Me.Icon = FR_Principal.Icon
  EDN_Archivo.Filter = "Archivo HTML (*.html) |*.html;*.htm| Todos los Archivos (*.*) |*.*"
  If guAccion <> etAcNuevo Then
    txtDescripcion.Text = guObjGeneral.QuitarParametro(gstrDatoIntermedio, "@")
    EDN_Archivo.NombreArchivo = gstrDatoIntermedio
  End If
End Sub
