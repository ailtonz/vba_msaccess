Attribute VB_Name = "BasEtiquetas"
Option Compare Database
Option Explicit

Private ContadorBrancos As Integer
Private ContadorCopias As Integer

Function GeraEtiquetas(R As Report, Optional NumEtiqCopias As Integer = 0, _
Optional NumEtiqBrancos As Integer = 0)

'===========================================================
' Rotina principal, � executada a partir da Propriedade
' Detalhe_Print da se��o  do Relat�rio.
' Esta rotina gerencia tanto as etiquetas a pular no in�cio
' do Relat�rio quanto o n�mero de c�pias a serem impressas.
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
'No evento AoAbrir do relat�rio chame esta rotina
'para zerar as vari�veis
ContadorBrancos = 0
ContadorCopias = 0
End Sub
   
   
Public Function ImprimirReport()
'para abrir a caixa de dialogo imprimir antes de
'mandar para a impressora
    DoCmd.RunCommand acCmdPrint
End Function

