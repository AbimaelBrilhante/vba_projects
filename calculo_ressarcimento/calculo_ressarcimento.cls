VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Calcula_ICMS()

Dim i As Long
Dim j As Long
Dim i_final As Long
Dim j_final As Long

'quantidade de linhas
i_final = Sheets(1).Cells(7, 3).Value
j_final = Sheets(1).Cells(6, 3).Value

'tempo de processamento
    Sheets(1).Cells(4, 3).Value = Time
    For i = 1 To i_final
    Sheets(1).Cells(3, 3).Value = i
    
    
    Product_OUT = Sheets(1).Cells(11 + i, 34).Value 'produto saida
    Data = Sheets(1).Cells(11 + i, 37).Value 'data saida
        
        For j = 1 To j_final
        Product_IN = Sheets(1).Cells(11 + j, 4).Value 'produto entrada
        ESTOQUE = Sheets(1).Cells(11 + j, 29).Value 'saldo entrada
            
            If (Product_IN = Product_OUT And ESTOQUE > 0 And Data >= Sheets(1).Cells(11 + j, 14).Value) Then 'data entrada
                    Sheets(1).Cells(11 + j, 29).Value = ESTOQUE - Sheets(1).Cells(11 + i, 38).Value 'saldo entrada e quantidade saida
                    
                    If Sheets(1).Cells(11 + j, 29).Value < 0 Then 'saldo entrada
                        Sheets(1).Cells(11 + j, 29).Value = Sheets(1).Cells(11 + j, 29).Value + Sheets(1).Cells(11 + i, 38).Value 'saldo entrada,saldo entrada e data saida
                        Sheets(1).Cells(11 + i, 48).Value = "AVALIAR" 'status
                        GoTo Proximo_Item
                    End If
                
                Sheets(1).Cells(11 + i, 44).Value = Sheets(1).Cells(11 + j, 27).Value 'calculo1
                Sheets(1).Cells(11 + i, 46).Value = Sheets(1).Cells(11 + j, 28).Value 'calculo 2
                Sheets(1).Cells(11 + i, 48).Value = "CREDITO APROPRIADO" 'status
                Sheets(1).Cells(11 + i, 49).Value = Sheets(1).Cells(11 + j, 13).Value 'nf entrada
                Sheets(1).Cells(11 + i, 52).Value = Sheets(1).Cells(11 + j, 14).Value 'data entrada
                Sheets(1).Cells(11 + i, 53).Value = Sheets(1).Cells(11 + j, 6).Value 'spool
                GoTo Saida
            
            End If
            Sheets(1).Cells(11 + i, 49).Value = "CREDITO NAO APROPRIADO"
Proximo_Item:
        Next j
Saida:
    Next i
Sheets(1).Cells(5, 3).Value = Time
End Sub


