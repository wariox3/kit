Attribute VB_Name = "ModFuncionesPropias"
Option Explicit
Public Function LlenarCombo(Combo As DataCombo, ListField As String, Tabla As String)
  AbrirRecorset rstUniversal, "select " & ListField & " from " & Tabla & " ORDER BY " & ListField, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Set Combo.RowSource = rstUniversal
  Combo.ListField = ListField
End Function
