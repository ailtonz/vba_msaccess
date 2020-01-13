Attribute VB_Name = "modInteracaoComSistema"
Option Compare Database

Public Function AbrirArquivo(sTitulo As String, sDecricao As String, sTipo As String, SelecaoMultipla As Boolean) As String
Dim fd As Office.FileDialog

'Di�logo de selecionar arquivo - Office
Set fd = Application.FileDialog(msoFileDialogFilePicker)

'T�tulo
fd.Title = sTitulo

'Filtros e descri��o dos mesmos
fd.Filters.Add sDecricao, sTipo

'Premiss�es de sela��o
fd.AllowMultiSelect = SelecaoMultipla

If fd.Show = -1 Then
    AbrirArquivo = fd.SelectedItems(1)
End If

End Function

Public Function Confirmar(sMensagem As String) As _
Boolean
'Faz uma pergunta ao usu�rio e retorma True se a
'resposta for SIM, e false se a resposta for N�O
Dim intResp As Integer

intResp = MsgBox(sMensagem, vbYesNo + vbQuestion, _
"Confirma��o")

If intResp = vbYes Then
    Confirmar = True
Else
    Confirmar = False
End If
End Function


