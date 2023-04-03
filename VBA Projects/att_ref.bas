sub att_ref()
        Application.ScreenUpdating = False
        If AbaExiste("Referências não localizadas") Then
            activeworkbook.Worksheets("Referências não localizadas").delete
        End If
    'bloco de definições (nomes das planilhas e abas)
        dim wbSKU, wbALPHA, wbMKT, wbGERAL as Workbook
        dim wsSKU, wsALPHA, wsMKT, wsGERAL, wssku2 as Worksheet
        set wbsku = Workbooks("sku completo")
        set wbalpha = Workbooks("Relatório interface alphaville")
        set wbmkt = Workbooks("Relatório interface market")
        set wbgeral = Workbooks("Relatório geral interface")
        set wssku = wbsku.Worksheets(1)
        set wsalpha = wbalpha.Worksheets(1)
        set wsmkt = wbmkt.Worksheets(1)
        set wsgeral = wbgeral.Worksheets(1)

    'verificação linha a linha da planilha geral, caso referência não seja encontrada na SKU, alimentar aba
        set wssku2 = wbsku.Worksheets.add
        wssku2.name = "Referências não localizadas"
        wssku2.Range("A1:M1").value = wsgeral.Range("A1:M1").value
        qtrefs = wsgeral.range("B1").End(xlDown).row
        qtrefs2 = wssku.range("A1").end(xlDown).row

        for lin_refs = 2 to qtrefs

            If wsgeral.range("D" & lin_refs) <> 0 or wsgeral.range("F" & lin_refs) <> 0 then

            refc = wsgeral.range("B" & lin_refs) & wsgeral.range("C" & lin_refs)
            
            For lin_refs2 = 2 to qtrefs2
                refc2 = wssku.range("A" & lin_refs2) & wssku.range("B" & lin_refs2)

                If refc = refc2 Then
                    goto proxrefgeral
                End If
            Next lin_refs2

            lin_atual = application.WorksheetFunction.CountA(wssku2.Range("A:A")) + 1
            wssku2.Range("A" & lin_atual & ":M" & lin_atual).value = wsgeral.Range("A" & lin_refs & ":M" & lin_refs).value

            End if 'verificar se referência possui valor
    proxrefgeral:
        next lin_refs
end sub
Function AbaExiste(Nome As String) As Boolean
        Dim Tmp As String
        AbaExiste = True
        On Error GoTo FIM
        Tmp = ThisWorkbook.Worksheets(Nome).Name
    FIM: If Err.Number = 9 Then AbaExiste = False
End Function
Sub att_valores()
    dim wbSKU, wbALPHA, wbMKT, wbGERAL as Workbook
    dim wsSKU, wsALPHA, wsMKT, wsGERAL, wssku2 as Worksheet
    set wbsku = Workbooks("sku completo")
    set wbalpha = Workbooks("Relatório interface alphaville")
    set wbmkt = Workbooks("Relatório interface market")
    set wbgeral = Workbooks("Relatório geral interface")
    set wssku = wbsku.Worksheets(1)
    set wsalpha = wbalpha.Worksheets(1)
    set wsmkt = wbmkt.Worksheets(1)
    set wsgeral = wbgeral.Worksheets(1)

    num_refs = wssku.Range("A1").End(xlDown).row
    For lin_ref = 2 to num_refs
        refc = wssku.Range("A" & lin_ref) & wssku.Range("B" & lin_ref)

        num_refsalpha = wsalpha.Range("A1").End(xlDown).Row
        For lin_refalpha = 2 to num_refsalpha
            refcalpha = wsalpha.Range("B" & num_refsalpha) & wsalpha.Range("C" & num_refsalpha)
            
            if refcalpha = refc then
                if wsalpha.Range("D" & lin_refalpha) <> 0 or wsalpha.Range("F" & lin_refalpha) <> 0 Then
                    qtdvalpha = wsalpha.Range("D" & lin_refalpha).value
                    qtdealpha = wsalpha.Range("F" & lin_refalpha).value
                    goto preenchimento
                else
                    qtdvalpha = 0
                    qtdealpha = 0
                End If 'verificar se a referência é diferente de 0
            else
                qtdvalpha = 0
                qtdealpha = 0
            End If 'verificar se a referência é igual à pesquisada
        Next

        num_refsmkt = wsmkt.Range("A1").End(xlDown).Row
        For lin_refmkt = 2 to num_refsmkt
            refcmkt = wsmkt.Range("B" & num_refsmkt) & wsmkt.Range("C" & num_refsmkt)
            
            if refcmkt = refc then
                if wsmkt.Range("D" & lin_refmkt) <> 0 or wsmkt.Range("F" & lin_refmkt) <> 0 Then
                    qtdvmkt = wsmkt.Range("D" & lin_refmkt).value
                    qtdemkt = wsmkt.Range("F" & lin_refmkt).value
                    goto preenchimento
                else
                    qtdvmkt = 0
                    qtdemkt = 0
                End If 'verificar se a referência é diferente de 0
            else
                qtdvmkt = 0
                qtdemkt = 0
            End If 'verificar se a referência é igual à pesquisada
        Next

        num_refsgeral = wsgeral.Range("A1").End(xlDown).Row
        For lin_refgeral = 2 to num_refsgeral
            refcgeral = wsgeral.Range("B" & num_refsgeral) & wsgeral.Range("C" & num_refsgeral)
            
            if refcgeral = refc then
                if wsgeral.Range("D" & lin_refgeral) <> 0 or wsgeral.Range("F" & lin_refgeral) <> 0 Then
                    qtdvgeral = wsgeral.Range("D" & lin_refgeral).value
                    qtdegeral = wsgeral.Range("F" & lin_refgeral).value
                    preco = wsgeral.Range("M" & lin_refgeral).value
                    goto preenchimento
                else
                    qtdvgeral = 0
                    qtdegeral = 0
                    preco = 0
                End If 'verificar se a referência é diferente de 0
            else
                qtdvgeral = 0
                qtdegeral = 0
                preco = 0
            End If 'verificar se a referência é igual à pesquisada
        Next

        preenchimento:
        wssku.Range("F" & lin_ref) = qtdealpha
        wssku.Range("G" & lin_ref) = qtdvalpha
        wssku.Range("H" & lin_ref) = qtdemkt
        wssku.Range("I" & lin_ref) = qtdvmkt
        wssku.Range("J" & lin_ref) = qtdegeral
        wssku.Range("K" & lin_ref) = qtdvgeral
        wssku.Range("M" & lin_ref) = preco

    Next lin_ref
End Sub