Attribute VB_Name = "CMD_APGAR"
Sub ApagarConteudoDaPasta()
    Dim fso As Object
    Dim pasta As Object
    Dim arquivos As Object
    Dim arquivo As Object

    ' Caminho da pasta que voc� deseja apagar
    Dim caminhoPasta As String
    caminhoPasta = "C:\Caminho\Para\Sua\Pasta\"

    ' Inicializar o objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Verificar se a pasta existe
    If fso.FolderExists(caminhoPasta) Then
        ' Obter refer�ncia para a pasta
        Set pasta = fso.GetFolder(caminhoPasta)

        ' Obter a cole��o de arquivos na pasta
        Set arquivos = pasta.Files

        ' Excluir cada arquivo na pasta
        For Each arquivo In arquivos
            arquivo.Delete
        Next arquivo

        ' Excluir subpastas (opcional)
        For Each subpasta In pasta.SubFolders
            subpasta.Delete
        Next subpasta

        MsgBox "Conte�do da pasta foi apagado com sucesso!", vbInformation
    Else
        MsgBox "A pasta n�o foi encontrada.", vbExclamation
    End If
End Sub


