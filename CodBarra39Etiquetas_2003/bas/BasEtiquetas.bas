Attribute VB_Name = "BasEtiquetas"
Option Compare Database
Option Explicit

Private ContadorBrancos As Integer
Private ContadorCopias As Integer

Function GeraEtiquetas(R As Report, Optional NumEtiqCopias As Integer = 0, _
Optional NumEtiqBrancos As Integer = 0)

'===========================================================
' Rotina principal, é executada a partir da Propriedade
' Detalhe_Print da seção  do Relatório.
' Esta rotina gerencia tanto as etiquetas a pular no início
' do Relatório quanto o número de cópias a serem impressas.
'===========================================================
  If ContadorBrancos < NumEtiqBrancos Then
        R.NextRecord = False
        R.PrintSection = False
        R.MoveLayout = True
        ContadorBrancos = ContadorBrancos + 1
  Else
        If ContadorCopias < NumEtiqCopias Then
           R.NextRecord = False
           R.PrintSection = True
           R.MoveLayout = True
           ContadorCopias = ContadorCopias + 1
        Else
           ContadorCopias = 0
        End If
  End If
End Function
Sub zeraVariaveis()
'No evento AoAbrir do relatório chame esta rotina
'para zerar as variáveis
ContadorBrancos = 0
ContadorCopias = 0
End Sub
   
   
Public Function ImprimirReport()
'para abrir a caixa de dialogo imprimir antes de
'mandar para a impressora
    DoCmd.RunCommand acCmdPrint
End Function

