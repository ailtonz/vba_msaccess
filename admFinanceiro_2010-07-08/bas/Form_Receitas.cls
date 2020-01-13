VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Receitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(Cancel As Integer)
    
    If Me.NewRecord Then
       Me.DataDeEmissao = Format(Now(), "dd/mm/yy")
       Me.codTipoMovimento = "Receita"
       Me.Status = "Aberto"
    End If
    
End Sub

Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    If EstaAberto("Pesquisar") Then Form_Pesquisar.lstCadastro.Requery
    DoCmd.Close

Exit_cmdSalvar_Click:
    Exit Sub

Err_cmdSalvar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdSalvar_Click
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    DoCmd.CancelEvent
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdFechar_Click

End Sub

Private Sub codCategoria_Click()
Dim SQL_Definicoes As String
Dim strCategoria As String

strCategoria = Me.codCategoria.Column(0)

SQL_Definicoes = "Select Distinct * from qryDefinicaoReceitas where Categoria = '" & Forms!Receitas.codCategoria.Column(0) & "'"

Me.codDefinicao.RowSource = SQL_Definicoes
        
End Sub

Private Sub codDefinicao_GotFocus()
    codCategoria_Click
End Sub

Private Sub cmdRepetir_Click()
Dim DB As DAO.Database
Dim rst As DAO.Recordset
Dim x As Integer
Dim VCTO As Integer
Dim MES As Integer


VCTO = Format(DataDeVencimento, "dd")
MES = Format(DataDeVencimento, "mm") + 1

'Salvar Registro
DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Set DB = CurrentDb()
'Abre a tabela, adiciona o novo item e atualiza a combo
Set rst = DB.OpenRecordset("Movimentos")
    
Dim Message, Title, Default, MyValue

Message = "Quantas vezes este cadastro deve repetir? "    ' Define o aviso.
Title = "Repetir cadastro"       ' Define o título.
Default = "1"    ' Define o padrão.

msgRepetirCadastro:
' Exibe a mensagem, o título e o valor padrão.
MyValue = InputBox(Message, Title, Default)

' Cancelar Processo
If MyValue = "" Then GoTo sair

' Verificar Integridade de informação (É numero?)
If Not IsNumeric(MyValue) Then GoTo msgRepetirCadastro
    
Me.Controle = "(" & 1 & "/" & MyValue & ")"

'Salvar Registro
DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

With rst
    
    For x = 2 To MyValue
        .AddNew
        !codMovimento = NovoCodigo("Movimentos", "codMovimento")
        !DataDeEmissao = Format(Now(), "dd/mm/yy")
        !DataDeVencimento = CalcularVencimento(VCTO, MES, Year(Now))
        !codTipoMovimento = Me.codTipoMovimento
        !DescricaoDoMovimento = Me.DescricaoDoMovimento
        !Categoria = Me.codCategoria
        !Definicao = Me.codDefinicao
        !Nome = Me.Nome
        !Controle = "(" & x & "/" & MyValue & ")"
        !ValorDoMovimento = Me.ValorDoMovimento
        !Banco = Me.Banco
        !Status = Me.Status
        !codRelacao = Me.Codigo
        .update
        MES = MES + 1
    Next x
    
End With

MsgBox "Operação realizada com sucesso!", vbOKOnly + vbInformation, Title

sair:

rst.Close
DB.Close

Set rst = Nothing
Set DB = Nothing

End Sub
