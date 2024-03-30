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

    depto = ws.cells(3,7)
    ano = ws.cells(2,7)
    mes = ws.cells(1,7)

End Sub