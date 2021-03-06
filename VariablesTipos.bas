Attribute VB_Name = "VariablesTipos"
'--> Variables / constantes / tipos globales
Option Explicit

Global Const gstrComillas As String = """"

Global Const ConstCadenaFecha As String = "dd-mm-yyyy"
Global Const ConstCadenaFechaAmericana As String = "mm-dd-yyyy"

Global Const ConstIconoSubir = "ICON_SUBIR"
Global Const ConstIconoNuevo = "ICON_NUEVO"
Global Const ConstIconoBorrar = "ICON_BORRAR"
Global Const ConstIconoBajar = "ICON_BAJAR"
Global Const ConstIconoInterrogacion = "ICON_INTERROGACION"
Global Const ConstIconoExclamacion = "ICON_EXCLAMACION"
Global Const ConstIconoInformacion = "ICON_INFORMACION"
Global Const ConstIconoCorteComunicacion = "ICON_CORTECOMUNICACION"
Global Const ConstIconoBaseDatos = "ICON_BASEDATOS"
Global Const ConstIconoDrag = "ICON_DRAG"
Global Const ConstIconoDragInvalido = "ICON_DRAGINVALIDO"
Global Const ConstIconoCalendario = "ICON_CALENDARIO"
Global Const ConstIconoLogoGiss = "LOGO_GISS"

Global guObjGeneral As New ClassRutinasGenerales 'Objeto a la clase con las rutinas de uso general

Public Enum etAmbito 'Ambito de las propiedades, rutinas, variables...
  etAmbPublico = 0 'Identificador de �mbito p�blico
  etAmbPrivado 'Identificador de �mbito privado
  etAmbGlobal 'Identificador de �mbito global
  etAmbProtegido 'Identificador de �mbito protegido
End Enum

Public Enum etAccion 'Tipos de acci�n
  etAcRealizado = 0 'Acci�n completa
  etAcCancelado 'Acci�n cancelada por el usuario
  etAcNuevo 'Nuevo
  etAcModificar 'Modificar
  etAcBorrar 'Borrar
End Enum

Global guAccion As etAccion 'Variable de intercambio que indica si deseamos hacer nuevo, modificar o borrar as� como si la acci�n est� completa o cancelada
Global gstrDatoIntermedio As String 'Variable para intercambio de par�metros entre formularios

Public Function getCadenaAmbito(ByVal Ambito As etAmbito) As String
'--> Obtiene la cadena con el �mbito de la variable o rutina (si le pasamos etAmbGlobal devuelve "Global")
'--> @param Ambito Tipo de ambito
'--> @return String cadena identificativa del tipo (global, p�blico, privado...)
  Select Case Ambito
    Case etAmbGlobal
      getCadenaAmbito = "Global"
    Case etAmbPublico
      getCadenaAmbito = "P�blico"
    Case etAmbPrivado
      getCadenaAmbito = "Privado"
    Case etAmbProtegido
      getCadenaAmbito = "Protegido"
    Case Else
      getCadenaAmbito = "DESCONOCIDO"
  End Select
End Function
