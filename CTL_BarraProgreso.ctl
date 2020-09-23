VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl BarraProgreso 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   ScaleHeight     =   360
   ScaleWidth      =   8400
   ToolboxBitmap   =   "CTL_BarraProgreso.ctx":0000
   Begin MSComctlLib.ProgressBar PR_Progreso 
      Height          =   252
      Left            =   216
      TabIndex        =   0
      Top             =   72
      Width           =   5436
      _ExtentX        =   9604
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "BarraProgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> Control de una barra de progreso
Option Explicit

Private Minimo As Long 'Valor mínimo de la barra
Private Maximo As Long 'Valor máximo de la barra

Public Sub Cambiar()
'--> Incrementa en uno la barra de progreso
  With PR_Progreso
    If .Value + 1 >= .Min And .Value + 1 <= .Max Then .Value = .Value + 1
  End With
End Sub

Public Sub Inicializar(ByVal Minimo As Long, ByVal Maximo As Long)
'--> Inicializa la barra de progreso.
  On Error Resume Next
  With PR_Progreso
    .Min = Minimo
    If Maximo > Minimo Then
      .Max = Maximo
    Else
      .Max = Minimo + 1
    End If
    .Value = Minimo
  End With
End Sub

Public Property Let Max(ByVal pValor As Integer)
Attribute Max.VB_Description = "Valor máximo de la barra de progreso"
  Inicializar Minimo, pValor
  PropertyChanged
End Property

Public Property Get Max() As Integer
  Max = Maximo
End Property

Public Property Let Min(ByVal pValor As Integer)
Attribute Min.VB_Description = "Valor mínimo de la barra de progreso"
  Inicializar pValor, Maximo
  PropertyChanged
End Property

Public Property Get Min() As Integer
  Min = Minimo
End Property

Private Sub UserControl_InitProperties()
  Max = 1
  Min = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Max = PropBag.ReadProperty("Max", 1)
  Min = PropBag.ReadProperty("Min", 0)
End Sub

Private Sub UserControl_Resize()
  PR_Progreso.Top = 0
  PR_Progreso.Left = 0
  PR_Progreso.Width = Width
  PR_Progreso.Height = Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Max", Maximo, 1
  PropBag.WriteProperty "Min", Minimo, 0
End Sub
