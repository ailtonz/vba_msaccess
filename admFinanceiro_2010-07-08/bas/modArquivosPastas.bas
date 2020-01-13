Attribute VB_Name = "modArquivosPastas"
Option Compare Database

Public Function VerificaExistenciaDeArquivo(Localizacao As String) As Boolean

If Dir(Localizacao, vbArchive) <> "" Then
    VerificaExistenciaDeArquivo = True
Else
    VerificaExistenciaDeArquivo = False
End If

End Function

Public Function getCaminho(arqCaminho As String) As String
Dim lin As String

Open arqCaminho For Input As #1

Line Input #1, lin
getCaminho = lin

Close #1

End Function

Public Function CriarPasta(sPasta As String) As String
'Cria pasta apartir da origem do sistema

Dim fPasta As New FileSystemObject
Dim MyApl As String

MyApl = Application.CurrentProject.Path

If Not fPasta.FolderExists(MyApl & "\" & sPasta) Then
   fPasta.CreateFolder (MyApl & "\" & sPasta)
End If

CriarPasta = MyApl & "\" & sPasta & "\"

End Function

Public Function getPath(sPathIn As String) As String
'Esta função irá retornar apenas o path de uma string que contenha o path e o nome do arquivo:
Dim i As Integer

  For i = Len(sPathIn) To 1 Step -1
     If InStr(":\", Mid$(sPathIn, i, 1)) Then Exit For
  Next

  getPath = Left$(sPathIn, i)

End Function

Public Function getFileName(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim i As Integer

  For i = Len(sFileIn) To 1 Step -1
     If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
  Next

  getFileName = Left(Mid$(sFileIn, i + 1, Len(sFileIn) - i), Len(Mid$(sFileIn, i + 1, Len(sFileIn) - i)) - 4)

End Function

Public Function getFileExt(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim i As Integer

  For i = Len(sFileIn) To 1 Step -1
     If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
  Next

  getFileExt = right(Mid$(sFileIn, i + 1, Len(sFileIn) - i), 4)

End Function

