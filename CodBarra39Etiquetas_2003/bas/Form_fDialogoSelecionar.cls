VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fDialogoSelecionar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Close()
'DoCmd.Quit
End Sub


Private Sub Limpa_Click()
Dim i As Integer
    For i = 0 To Me.ListaProdutos.ListCount - 1
    Me.ListaProdutos.Selected(i) = False
    Next i
End Sub

Private Sub OK_Click()
'Verifica se algum item foi selecionado
'Duas Formas
'If Me.ListaProdutos.ListIndex = -1 Then Exit Sub
If Me.ListaProdutos.ItemsSelected.Count = 0 Then
    DoCmd.Beep
    MsgBox "Nenhum registro selecionado!", vbInformation, "ATENÇÃO"
Else
    Call SelecaoMultiplaR(Me.ListaProdutos, "CódigoDoProduto", "rpt_BarCode")
    Me.Visible = False
End If
End Sub

Private Sub Todos_Click()
Dim i As Integer
    For i = 0 To Me.ListaProdutos.ListCount - 1
    Me.ListaProdutos.Selected(i) = True
    Next i
End Sub
