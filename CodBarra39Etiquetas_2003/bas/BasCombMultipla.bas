Attribute VB_Name = "BasCombMultipla"
Option Compare Database
Option Explicit

Public Sub SelecaoMultiplaR(ListBox As Control, sCampo As String, _
NomeRel As String)
'Para abrir Formulários ou Relatorios com valores selecionados
'Autor: Carlos Moura em 10/08/98 e-mail: crpmoura@ig.com.br
'Num formulario com uma caixa de seleção múltipla

Dim varItem As Variant, strList As String, strWhere As String

With ListBox
    For Each varItem In .ItemsSelected
    'Aqui você concatena do jeito que quiser
    'Para valor campo string & "'" & ",'"
    'Para valor campo numerico & ","
    
    strList = strList & .Column(0, varItem) & ","
    Next varItem
End With
    'Para valor campo string ('" & strList & "')
    'Para valor campo numerico (" & strList & ")
    strWhere = sCampo & " In (" & strList & ")"
    DoCmd.OpenReport NomeRel, acPreview, , strWhere
End Sub



