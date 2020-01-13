VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rpt_BarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub CabeçalhoDoRelatório_Print(Cancel As Integer, PrintCount As Integer)
Call zeraVariaveis 'para imprimir em branco
End Sub

Private Sub Detail1_Print(Cancel As Integer, PrintCount As Integer)
Dim Result As Long 'Variavel para chamar a função codigo de barras
Dim numCopia As Integer, pula As Integer

On Error Resume Next
    'Chame a função assim para criar os códigos de Barra
    'Nome do controle caixa de texto para conter o desenho - não acoplada
    'e o relatório atual
    Result = MD_Barcode39(CodBarra, Me)
    
    'código para o gerador de etiquetas
    numCopia = Forms!fDialogoSelecionar!txtCopias
    pula = Forms!fDialogoSelecionar!txtBranco
    Call GeraEtiquetas(Me, numCopia, pula)

End Sub

Private Sub Report_Close()
On Error Resume Next
Forms!fDialogoSelecionar.Visible = True
End Sub

Private Sub Report_Open(Cancel As Integer)
Call zeraVariaveis 'para imprimir em branco
DoCmd.Maximize
End Sub
