Sub criar()
    Dim caminho, pDados, caminhoCompleto, prof, comp, curso, modulo, pMencoes, aluno As String
    Dim c, cprof, cc, caluno, csituacao, linhasituacao, colunasituacao

    ' Desabilitar Atualização da tela deixa o processo mais rápido 
    ' e a tela não fica piscando com a troca constante entra planilhas
    Application.ScreenUpdating = False
    pDados = "DADOS PARA PLANILHAS.xlsm"
    modulo = Range("A4")
    curso = Range("B4")
    For c = 6 To Range("D2")
        pMencoes = Range("A" & c)
        caminho = ActiveWorkbook.path & "\"
        FileCopy Source:=caminho & "MODELO.xlsx", Destination:=caminho & "MODELOCOPY.xlsx"
        caminhoCompleto = caminho & Range("A" & c) & ".xlsx"
        Name caminho & "MODELOCOPY.xlsx" As caminhoCompleto

        cprof = c - 2
        prof = Range("J" & cprof)
        comp = Range("I" & cprof)


        Workbooks.Open (caminhoCompleto)
        Worksheets("Frente").Activate

        Range("N3") = prof
        Range("I4") = curso
        Range("P4") = modulo
        Range("Q3") = comp

        For cc = 1 To 50
            Workbooks(pDados).Activate
            Worksheets("Plan1").Activate
            caluno = cc + 2
            aluno = Range("F" & caluno)
            colunasituacao = cc + 11
            linhasituacao = c - 2
            situacao = Cells(linhasituacao, colunasituacao)
            Workbooks(pMencoes).Activate
            Worksheets("Frente").Activate
            caluno = cc + 5
            Range("B" & caluno) = aluno
            Range("S" & caluno) = situacao
            Range("T" & caluno) = situacao
            Range("U" & caluno) = situacao
            Range("V" & caluno) = situacao
            Range("W" & caluno) = situacao
            Range("X" & caluno) = situacao
        Next cc


        Workbooks(pMencoes).Close savechanges:=True

        Workbooks(pDados).Activate


    Next c
    Application.ScreenUpdating = True

End Sub