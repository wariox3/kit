Attribute VB_Name = "ModFuncionesBasicas"
Option Explicit
Sub EnfocarT(CajaTexto As TextBox)
  CajaTexto.SelStart = 0
  CajaTexto.SelLength = Len(CajaTexto)
End Sub
Public Sub EnfocarM(Caja As MaskEdBox)
  Caja.SelStart = 0
  Caja.SelLength = Len(Caja)
End Sub
Sub IniProg(Fin)
On Error GoTo MiError
  Principal.PgsPrincipal.Width = Principal.Picture1.Width
  Principal.PgsPrincipal.Min = 1
  Principal.PgsPrincipal.Max = Fin
  Principal.TxtMensaje.Visible = False
MiError:
End Sub
Sub Prog(Valor)
On Error GoTo MiError
  Principal.PgsPrincipal.Value = Valor
MiError:
End Sub
Sub FinProg()
  Principal.PgsPrincipal.Width = 0
  Principal.TxtMensaje.Visible = True
End Sub


