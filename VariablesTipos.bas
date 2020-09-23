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
  etAmbPublico = 0 'Identificador de ámbito público
  etAmbPrivado 'Identificador de ámbito privado
  etAmbGlobal 'Identificador de ámbito global
  etAmbProtegido 'Identificador de ámbito protegido
End Enum

Public Enum etAccion 'Tipos de acción
  etAcRealizado = 0 'Acción completa
  etAcCancelado 'Acción cancelada por el usuario
  etAcNuevo 'Nuevo
  etAcModificar 'Modificar
  etAcBorrar 'Borrar
End Enum

Global guAccion As etAccion 'Variable de intercambio que indica si deseamos hacer nuevo, modificar o borrar así como si la acción está completa o cancelada
Global gstrDatoIntermedio As String 'Variable para intercambio de parámetros entre formularios

Public Function getCadenaAmbito(ByVal Ambito As etAmbito) As String
'--> Obtiene la cadena con el ámbito de la variable o rutina (si le pasamos etAmbGlobal devuelve "Global")
'--> @param Ambito Tipo de ambito
'--> @return String cadena identificativa del tipo (global, público, privado...)
  Select Case Ambito
    Case etAmbGlobal
      getCadenaAmbito = "Global"
    Case etAmbPublico
      getCadenaAmbito = "Público"
    Case etAmbPrivado
      getCadenaAmbito = "Privado"
    Case etAmbProtegido
      getCadenaAmbito = "Protegido"
    Case Else
      getCadenaAmbito = "DESCONOCIDO"
  End Select
End Function
