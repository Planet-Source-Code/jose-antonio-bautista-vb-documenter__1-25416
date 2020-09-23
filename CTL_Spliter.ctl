VERSION 5.00
Begin VB.UserControl Spliter 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6420
   ClipControls    =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6420
   ToolboxBitmap   =   "CTL_Spliter.ctx":0000
   Begin VB.PictureBox PI_Spliter 
      BorderStyle     =   0  'None
      Height          =   4188
      Left            =   1320
      MouseIcon       =   "CTL_Spliter.ctx":0182
      MousePointer    =   99  'Custom
      ScaleHeight     =   4185
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   72
      Width           =   60
   End
   Begin VB.PictureBox IM_Spliter 
      BackColor       =   &H80000008&
      Height          =   3396
      Left            =   2664
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3390
      ScaleWidth      =   30
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   36
   End
End
Attribute VB_Name = "Spliter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> Control de usuario para el manejo de las barras que permiten establecer _
     el tamaño de determinados elementos de la ventana (Spliter Vertical)
Option Explicit

Public Event Redimensionar(ByVal x As Integer)

Private Redimensionando As Boolean
Private SpliterMinimo As Integer

Property Let Posicion(ByVal x As Integer)
'--> Consigue la posición inicial de la barra de Spliter
  PI_Spliter.Left = x
End Property

Property Get Posicion() As Integer
'--> Consigue la posición actual de la barra de Spliter
  Posicion = PI_Spliter.Left
End Property

Property Let Minimo(ByVal SplMinimo As Integer)
'--> Asigna la posición mínima de la barra de Spliter
  SpliterMinimo = SplMinimo
  PropertyChanged
End Property

Property Get Minimo() As Integer
'--> Devuelve la posición mínima de la barra de Spliter
  Minimo = SpliterMinimo
End Property

Public Sub Redimensionar(Optional ByVal InterWidth As Integer = -1, Optional ByVal InterHeight As Integer = -1)
'--> Redimensiona el tamaño del control
  'Tener en cuenta que si está dentro de un control de usuario no tiene parent
  On Error Resume Next
  If InterHeight = -1 Then
    Height = Parent.Height
  Else
    Height = InterHeight
  End If
  If InterWidth = -1 Then
    Width = Parent.Width
  Else
    Width = InterWidth
  End If
  With PI_Spliter
    .Top = 0
    .Height = Height
  End With
End Sub

Private Sub PI_Spliter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'--> Si se pulsa el botón sobre el Spliter comienza a redimensionar
  Redimensionando = (Button = vbLeftButton)
  If Redimensionando Then
    With PI_Spliter
      IM_Spliter.Move .Left, .Top + 20, .Width / 2, .Height - 40
    End With
  End If
End Sub

Private Sub PI_Spliter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'--> Si se mueve el Spliter mientras está pulsado se realiza el movimiento
Dim Posicion As Single

  If Redimensionando And Button = vbLeftButton Then
    Posicion = x + PI_Spliter.Left
    If Posicion < SpliterMinimo Then
      IM_Spliter.Left = SpliterMinimo
    ElseIf Posicion > Width - SpliterMinimo Then
      IM_Spliter.Left = Width - SpliterMinimo
    Else
      IM_Spliter.Left = Posicion
    End If
  End If
End Sub

Private Sub PI_Spliter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'--> Cuando se suelta el ratón se lanza un evento <B> Redimensionar </B> al propietario del control
  If Redimensionando Then
    PI_Spliter.Left = IM_Spliter.Left
    RaiseEvent Redimensionar(IM_Spliter.Left)
    Redimensionando = False
    DoEvents
  End If
End Sub

Private Sub UserControl_Initialize()
  Redimensionando = False
  'Redimensionar
  PI_Spliter.Left = Width / 2
End Sub

Private Sub UserControl_InitProperties()
  Minimo = 1500
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Minimo = PropBag.ReadProperty("Minimo", 1500)
End Sub

Private Sub UserControl_Show()
  Redimensionar
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Minimo", SpliterMinimo, 1500
End Sub
