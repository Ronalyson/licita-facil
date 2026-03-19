' ============================================================
' Script: column-percentage-increase
' Descrição:
' Aumenta os valores de uma coluna inteira com base em uma
' porcentagem informada pelo usuário.
'
' Funcionamento:
' - Solicita a coluna que será alterada
' - Solicita a porcentagem de aumento
' - Percorre automaticamente as células da coluna escolhida
' - Aplica a alteração apenas em células numéricas,
'   não vazias e sem fórmulas
'
' Exemplo:
' Valor 100 com aumento de 15% = 115
'
' Objetivo:
' Facilitar ajustes em massa de valores em planilhas,
' evitando edições manuais repetitivas.
' ============================================================

Sub AumentarPorcentagemColuna()

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim cel As Range
    Dim coluna As String
    Dim porcentagem As Double
    Dim fator As Double
    
    Set ws = ActiveSheet
    
    ' Pergunta a coluna
    coluna = InputBox("Digite a coluna que deseja alterar (ex: A, B, C):")
    
    If coluna = "" Then Exit Sub
    
    ' Pergunta a porcentagem
    porcentagem = InputBox("Digite a porcentagem que deseja aumentar (ex: 15 para 15%)")
    
    fator = 1 + (porcentagem / 100)
    
    ' Última linha da coluna
    ultimaLinha = ws.Cells(ws.Rows.Count, coluna).End(xlUp).Row
    
    ' Percorre as células
    For Each cel In ws.Range(coluna & "1:" & coluna & ultimaLinha)
        
        If Not cel.HasFormula And IsNumeric(cel.Value) And cel.Value <> "" Then
            cel.Value = cel.Value * fator
        End If
        
    Next cel
    
    MsgBox "Aumento de " & porcentagem & "% aplicado na coluna " & coluna & "!", vbInformation

End Sub