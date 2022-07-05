Attribute VB_Name = "Modulo1"
Option Explicit


    ' Cria uma nova instancia do FileSystemObject
    Private fso As New FileSystemObject

    ' Declaracao de variaveis
    Private sFolder As Scripting.Folder
    Private myFile As Scripting.File
    Private myFolder As Scripting.Folder
    Private pos As Long

    ' --------------------------------------------------------------------
    ' Sub auxiliar que lista os ficheiros nas sub-pastas
    ' --------------------------------------------------------------------
    Sub ShowSubFolderFiles(ByVal folderName As String)
    
  

        ' Ignora erros em pastas protegidas, pastas de sistema, etc
        On Error Resume Next

        Set sFolder = fso.GetFolder(folderName)

        ' Ciclo em todas as pastas
        For Each myFolder In sFolder.SubFolders

        ' Ciclo em todos os ficheiros
            For Each myFile In myFolder.Files
                Sheets("Script").Cells(pos + 1, 8).Value = myFile.Name
                Sheets("Script").Cells(pos + 1, 9).Value = myFile.DateCreated
                Sheets("Script").Cells(pos + 1, 10).Value = myFile.DateLastModified
            
                pos = pos + 1
            Next

            ' Recursividade: Caso a pasta tenha sub-pastas chama novamente
            ' o mesmo codigo - Sub ShowSubFolderFiles()
            If myFolder.SubFolders.Count > 0 Then
                ShowSubFolderFiles myFolder.Path
            End If

        Next

    End Sub

 

    ' --------------------------------------------------------------------
    ' Sub principal que inicia o processo de listagem de
    ' todos os ficheiros, com base numa pasta inicial
    ' --------------------------------------------------------------------
    Sub ShowFolderFiles()
  
    
        Dim initialFolder As String

        ' Ignora erros em pastas protegidas, pastas de sistema, etc
        On Error Resume Next

        ' Pasta inicial e linha da folha de calculo onde sera
        ' iniciada a escrita dos ficheiros encontrados
        initialFolder = Sheets("Script").Cells(1, 12) 'PASTA DE ORIGEM
        pos = 1
        


        ' Verifica se a pasta indicada existe
        If Not fso.FolderExists(initialFolder) Then
            MsgBox "A pasta " & initialFolder & " nao existe!", vbCritical
            Exit Sub
        End If

        ' Define a pasta inicial
        Set sFolder = fso.GetFolder(initialFolder)

        ' Mostra os ficheiros na pasta inicial
        ' incrementado a posicao (linha) onde escreve
        For Each myFile In sFolder.Files
            
            Sheets("Script").Cells(pos + 1, 8).Value = myFile.Name
            Sheets("Script").Cells(pos + 1, 9).Value = myFile.DateCreated
            Sheets("Script").Cells(pos + 1, 10).Value = myFile.DateLastModified
            pos = pos + 1
        Next

        ' Mostra os ficheiros nas sub-pastas
        Call ShowSubFolderFiles(sFolder.Path)

        ' Limpa as variaveis da memoria
        Set fso = Nothing
        Set myFolder = Nothing
        Set myFile = Nothing

  ActiveWorkbook.Save
  Application.DisplayFullScreen = False
  ActiveWindow.DisplayWorkbookTabs = True
  ActiveWindow.DisplayHorizontalScrollBar = True

End Sub




