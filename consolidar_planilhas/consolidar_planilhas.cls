VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Importar_XLS()

Dim sPath As String, sName As String, fName As String
Dim r As Long, rTemp As Long
Dim shPadrao As Worksheet

'Para a macro executar mais rápido!
With Application
.ScreenUpdating = False
.DisplayAlerts = False
End With

'A planilha onde serão colados os dados
Set shPadrao = Sheets("MACRO CONSOLIDAR")

'O caminho onde as planilhas que serão lidas estao
sPath = "C:\Users\abimaelsoares\Desktop\Nova pasta\"

'Descubro o nome do primeiro arquivo a ser aberto
sName = Dir(sPath & "*.xlsm*")

'Faço o loop que le todos os arquivos
Do While sName <> ""

'Acha a ultima linha utilizada na planilha onde serao colados os dados
r = shPadrao.Cells(Rows.Count, "A").End(xlUp).Row


'O caminho + o nome do arquivo a ser aberto
fName = sPath & sName

'Abro o workbook a ser lido
Workbooks.Open Filename:=fName, UpdateLinks:=False

'Descubro sua quantas linhas ele possui
rTemp = ActiveWorkbook.ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row


'Colo na planilha principal
ActiveWorkbook.ActiveSheet.Range("A7:O" & rTemp).Copy shPadrao.Range("A" & r + 1)

'Fecho o arquivo já lido
ActiveWorkbook.Close SaveChanges:=False

ScapeB:

'Atualizo a variavel com funcao DIR() que acha o proximo arquivo nao processado
sName = Dir()

Loop

On Error GoTo 0

With Application
.ScreenUpdating = True
.DisplayAlerts = True
End With

End Sub

Sub Limpar()
    Sheets(1).Range("A2:O30000").Clear
    
End Sub

