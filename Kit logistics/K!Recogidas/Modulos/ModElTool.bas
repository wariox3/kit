Attribute VB_Name = "ModElToolPermisos"
Option Explicit
Sub BotTool(Inicio As Byte, Fin As Byte, ElTool As Toolbar, Bloqueo As Boolean)
  If Bloqueo = True Then
    For II = Inicio To Fin
      If II <> 7 Then
        ElTool.Buttons.Item(II).Enabled = False
      End If
    Next
    ElTool.Buttons.Item(4).Enabled = True
    ElTool.Buttons.Item(7).Enabled = True
    ElTool.Buttons.Item(18).Enabled = False
    ElTool.Buttons.Item(19).Enabled = False
  Else
    For II = Inicio To Fin
      If II <> 7 Then
        ElTool.Buttons.Item(II).Enabled = True
      End If
    Next
    ElTool.Buttons.Item(4).Enabled = False
    ElTool.Buttons.Item(7).Enabled = False
    ElTool.Buttons.Item(18).Enabled = True
    ElTool.Buttons.Item(19).Enabled = True
  End If
End Sub
Public Sub LaTecla(LaTec As KeyCodeConstants, ElTool As Toolbar)
  Select Case LaTec
    Case vbKeyF3 ' Eliminar
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (6)
    Case vbKeyF4 ' Cancelar
      If ElTool.Buttons.Item(4).Enabled = True Then Screen.ActiveForm.AccionTool (7)
    Case vbKeyF5 ' Primero
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (11)
    Case vbKeyF6 ' Anterior
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (12)
    Case vbKeyF7 ' Siguiente
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (13)
    Case vbKeyF8 ' Ultimo
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (14)
    Case vbKeyF9 ' Nuevo
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (3)
    Case vbKeyF11 ' Guardar
      If ElTool.Buttons.Item(4).Enabled = True Then Screen.ActiveForm.AccionTool (4)
    Case vbKeyF10 ' Editar
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (5)
    Case vbKeyF12 ' Cerrar
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (16)
    Case vbKeyHome
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (9)
    Case vbKeyEnd ' Imprimir
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (18)
    Case vbKeyPause  ' Cargar informacion adicional
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (19)
  End Select
End Sub
'Public Function ComPermiso(Formulario As Byte, Permiso As Byte) As Boolean
'On Error GoTo IntNoExite
'  Select Case Permiso
'    Case 1
'      If Permisos(Formulario).Ingreso = True Then ComPermiso = True Else ComPermiso = False: MsgBox "No tiene permiso para ingresar a esta opcion", vbCritical, "Permiso restringido"
'    Case 2
'      If Permisos(Formulario).Nuevo = True Then ComPermiso = True Else ComPermiso = False: MsgBox "No tiene permiso para agreagar un nuevo registro", vbCritical, "Permiso restringido"
'    Case 3
'      If Permisos(Formulario).Editar = True Then ComPermiso = True Else ComPermiso = False: MsgBox "No tiene permiso para editar registros en este formulario", vbCritical, "Permiso restringido"
'    Case 4
'      If Permisos(Formulario).Eliminar = True Then ComPermiso = True Else ComPermiso = False: MsgBox "No tiene permiso para eliminar", vbCritical, "Permiso restringido"
'  End Select
'IntNoExite:
'  If Err.Number = 9 Then MsgBox "No ha sido habilitado para la accion que desea ejecutar, consulte al administrador", vbCritical, "No esta habilitado"
'End Function


