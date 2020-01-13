VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmExemplo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub txtNome_Change()
'By JPaulo ® Maximo Access
Dim i As Integer, intSomaCarateres As Integer

'Limita o campo a 20 carateres
    If Len(Me.txtNome.Text) > 20 Then
'Se chegar aos 20 carateres, mantem apenas os 20 digitados à esquerda
        Me.txtNome.Text = Left(Me.txtNome.Text, 20)
    Me.txtNome.SelStart = 20
'Informa o usuário, através de mensagem
MsgBox "Tamanho máximo do campo, excedido...", vbCritical
End If
'Inicia o contador a zero
    intSomaCarateres = 0
For i = 1 To Me.txtNome.SelStart
        intSomaCarateres = i
    Next i
'Coloca no campo do total de carateres, um a um
Me.txtTotalCarateres.Value = intSomaCarateres
'Se o usuário apagar os carateres um a um e chegar a zero, limpa a caixa
    If Me.txtTotalCarateres.Value = 0 Then
        Me.txtTotalCarateres.Value = ""
End If
End Sub


