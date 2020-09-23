VERSION 5.00
Begin VB.Form FR_AcercaDe 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acerca de MiApli"
   ClientHeight    =   4095
   ClientLeft      =   2340
   ClientTop       =   1890
   ClientWidth     =   8475
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3120
      TabIndex        =   0
      Top             =   3600
      Width           =   1245
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Información..."
      Height          =   345
      Left            =   4830
      TabIndex        =   1
      Top             =   3600
      Width           =   1245
   End
   Begin VB.Image IMG_Logo 
      BorderStyle     =   1  'Fixed Single
      Height          =   2700
      Left            =   96
      Top             =   60
      Width           =   3240
   End
   Begin VB.Label lblDescription 
      Caption         =   "Aplicación para el manejo de las bases de datos de G.I.S.S."
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   4380
      TabIndex        =   2
      Top             =   1620
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Título de la aplicación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   480
      Left            =   4380
      TabIndex        =   4
      Top             =   270
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versión"
      Height          =   225
      Left            =   4380
      TabIndex        =   5
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Versión de Desarrollo"
      ForeColor       =   &H00000000&
      Height          =   588
      Left            =   4368
      TabIndex        =   3
      Top             =   2940
      Width           =   3828
   End
End
Attribute VB_Name = "FR_AcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Ventana con el Acerca De... de la aplicación
'--> Se ha recogido exactamente de los formularios de ejemplo de VB 5.0.
Option Explicit

' Opciones de seguridad de claves del Registro...
Private Const READ_CONTROL = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos principales de claves del Registro...
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const ERROR_SUCCESS = 0
Private Const REG_SZ = 1                         ' Cadena Unicode terminada en Null
Private Const REG_DWORD = 4                      ' Número de 32 bits

Private Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Private Const gREGVALSYSINFOLOC = "MSINFO"
Private Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Private Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdSysInfo_Click()
'--> Muestra la información del sistema
Dim rc As Long
Dim SysInfoPath As String
  
  On Error GoTo SysInfoErr
  ' Prueba a obtener del Registro la información del sistema sobre el nombre y la ruta del programa...
  If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
  ' Prueba a obtener del Registro la información del sistema sobre la ruta del programa...
  ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
    ' Comprueba la existencia de una versión conocida de un archivo de 32 bits
    If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
      SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
    ' Error - Imposible encontrar el archivo...
    Else
      GoTo SysInfoErr
    End If
  ' Error - Imposible encontrar la entrada de Registro...
  Else
    GoTo SysInfoErr
  End If
  Call Shell(SysInfoPath, vbNormalFocus)
  Exit Sub
  
SysInfoErr:
  guObjGeneral.MensajeError "No está disponible la información de sistema.", mErrError
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
'--> Obtiene un valor de una clave del registro
Dim i As Long                                           ' Contador de bucle
Dim rc As Long                                          ' Código de retorno
Dim hKey As Long                                        ' Controlador a una clave de Registro abierta
Dim hDepth As Long                                      '
Dim KeyValType As Long                                  ' Tipo de dato de una clave de Registro
Dim tmpVal As String                                    ' Almacén temporal de una valor de clave de Registro
Dim KeyValSize As Long                                  ' Tamaño de la variable de la clave de Registro
  '------------------------------------------------------------
  ' Abre la clave de Registro en la raíz {HKEY_LOCAL_MACHINE...}
  '------------------------------------------------------------
  rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abre la clave de Registro
  If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Trata el error...
  tmpVal = String$(1024, 0)                               ' Asigna espacio para la variable
  KeyValSize = 1024                                       ' Marca el tamaño de la variable
  '------------------------------------------------------------
  ' Recupera valores de claves de Registro...
  '------------------------------------------------------------
  rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                       KeyValType, tmpVal, KeyValSize)    ' Obtiene o crea un valor de clave
  If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Trata el error
  If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agrega una cadena terminada en Null...
      tmpVal = Left(tmpVal, KeyValSize - 1)               ' Se encontró Null, se extrae de la cadena
  Else                                                    ' WinNT no tiene una cadena terminada en Null...
      tmpVal = Left(tmpVal, KeyValSize)                   ' No se encontró Null, sólo se extrae la cadena
  End If
  '------------------------------------------------------------
  ' Determina el tipo de valor de la clave para conversión...
  '------------------------------------------------------------
  Select Case KeyValType                                  ' Busca tipos de datos...
  Case REG_SZ                                             ' Tipo de dato de la cadena de la clave de Registro
    KeyVal = tmpVal                                     ' Copia el valor de la cadena
  Case REG_DWORD                                          ' El tipo de dato de la cadena de la clave es Double Word
    For i = Len(tmpVal) To 1 Step -1                    ' Convierte cada byte
      KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Genera el valor carácter a carácter
    Next
    KeyVal = Format$("&h" + KeyVal)                     ' Convierte Double Word a String
  End Select
  GetKeyValue = True                                      ' Vuelve con éxito
  rc = RegCloseKey(hKey)                                  ' Cierra la clave de Registro
  Exit Function                                           ' Salir
  
GetKeyError:      ' Restaurar después de que ocurra un error...
  KeyVal = ""                                             ' Establece el valor de retorno para una cadena vacía
  GetKeyValue = False                                     ' Devuelve un error
  rc = RegCloseKey(hKey)                                  ' Cierra la clave de Registro
End Function

Private Sub Form_Load()
  IMG_Logo.Picture = LoadResPicture(ConstIconoLogoGiss, vbResBitmap)
End Sub
