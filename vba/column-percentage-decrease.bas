' ============================================================
' Script: column-percentage-decrease
' Descrição:
' Diminui os valores de uma coluna inteira com base em uma
' porcentagem informada pelo usuário, com opção de arredondamento.
'
' Funcionamento:
' - Solicita a coluna que será alterada
' - Solicita a porcentagem de redução
' - Permite escolher o tipo de arredondamento:
'   C = para cima
'   B = para baixo
'   N = sem arredondamento
' - Percorre automaticamente as células da coluna escolhida
' - Aplica a alteração apenas em células numéricas,
'   não vazias e sem fórmulas
'
' Exemplo:
' Valor 3 com redução de 50% = 1,5
' - C (cima)  = 2
' - B (baixo) = 1
' - N (nenhum)= 1,5
'
' Objetivo:
' Facilitar reduções em massa de valores em planilhas,
' com controle sobre o arredondamento final.
' ============================================================

Sub DiminuirPorcentagemColuna()

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim cel As Range
    Dim coluna As String
    Dim porcentagem As Double
    Dim fator As Double
    Dim opcao As String
    Dim resultado As Double
    
    Set ws = ActiveSheet
    
    ' Pergunta a coluna
    coluna = InputBox("Digite a coluna que deseja alterar (ex: A, B, C):")
    If coluna = "" Then Exit Sub
    
    ' Pergunta a porcentagem
    porcentagem = InputBox("Digite a porcentagem que deseja DIMINUIR (ex: 50 para metade):")
    
    fator = 1 - (porcentagem / 100)
    
    ' Pergunta tipo de arredondamento
    opcao = InputBox("Escolha o tipo de arredondamento:" & vbCrLf & _
                     "C = arredondar para CIMA" & vbCrLf & _
                     "B = arredondar para BAIXO" & vbCrLf & _
                     "N = não arredondar")
    
    ' Última linha usada
    ultimaLinha = ws.Cells(ws.Rows.Count, coluna).End(xlUp).Row
    
    For Each cel In ws.Range(coluna & "1:" & coluna & ultimaLinha)
    
        If Not cel.HasFormula And IsNumeric(cel.Value) And cel.Value <> "" Then
            
            resultado = cel.Value * fator
            
            Select Case UCase(opcao)
                Case "C"
                    resultado = WorksheetFunction.RoundUp(resultado, 0)
                Case "B"
                    resultado = WorksheetFunction.RoundDown(resultado, 0)
            End Select
            
            cel.Value = resultado
            
        End If
        
    Next cel
    
    MsgBox "Redução de " & porcentagem & "% aplicada na coluna " & coluna & "!", vbInformation

End Sub