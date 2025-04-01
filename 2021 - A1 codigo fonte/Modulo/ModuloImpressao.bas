Attribute VB_Name = "ModuloImpressao"
Option Explicit

Public Sub ImpPV(IdReg As Integer)
'Pre-Venda
    Dim sSQL        As String
    Dim Rst1        As Recordset
    Dim Rst2        As Recordset
    If IdReg = 0 Then Exit Sub
    sSQL = "SELECT * FROM FaturamentoPV WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & IdReg
    Set Rst1 = RegistroBuscar(sSQL)
    'Set rptPreVenda.DataSource = Rst.DataSource
    
    sSQL = "SELECT * FROM FaturamentoPVItens WHERE ID_Empresa = " & ID_Empresa & " AND IDPV = " & IdReg
    'sSQL = "SELECT FaturamentoPVItens.*, FaturamentoPVItens.SubTotal + FaturamentoPVItens.VlIPI AS VlProdBruto FROM FaturamentoPVItens WHERE IDPV = " & IdReg
    
    Set Rst2 = RegistroBuscar(sSQL)
    Set rptPreVenda.DataSource = Rst2.DataSource
    Dim tpDoc As String
    
    If Len(cNull(Rst1.Fields("status"))) = 0 Then
            tpDoc = "PRE-VENDA"
        Else
            Select Case Rst1.Fields("status")
                Case "1"
                    tpDoc = "Orçamento"
                Case "2"
                    tpDoc = "Pedido"
                Case Else
                    tpDoc = "PRE-VENDA"
            End Select
    End If
            
    
    
    rptPreVenda.Title = "Documento_" & Left(String(10, "0"), 10 - Len(IdReg)) & IdReg & "-" & Format(Rst1.Fields("Emissao"), "DD_MM_YYYY")
    DoEvents
    '**************************************************************************************
    rptPreVenda.Sections("Section4").Controls.Item("LblTitulo").Caption = tpDoc
    
    rptPreVenda.Sections("Section4").Controls.Item("LblEmissao").Caption = Rst1.Fields("Emissao") 'dtpEmissao.Value
    rptPreVenda.Sections("Section4").Controls.Item("LblNumero").Caption = Left(String(5, "0"), 5 - Len(IdReg)) & IdReg
    '**************************************************************************************
    rptPreVenda.Sections("Section2").Controls.Item("LblNome").Caption = Rst1.Fields("Cliente")
    rptPreVenda.Sections("Section2").Controls.Item("LblFone").Caption = IIf(IsNull(Rst1.Fields("Tel")), "", Rst1.Fields("Tel"))
    rptPreVenda.Sections("Section2").Controls.Item("LblEndCliente").Caption = PgDadosCliente(Rst1.Fields("IdCliente")).Lgr & " " & PgDadosCliente(Rst1.Fields("IdCliente")).Nro & " " & PgDadosCliente(Rst1.Fields("IdCliente")).Cpl
    rptPreVenda.Sections("Section2").Controls.Item("LblCNPJ").Caption = PgDadosCliente(Rst1.Fields("IdCliente")).Doc 'IIf(PgDadosCliente(Rst1.Fields("IdCliente")).pessoa = "Fisica", Format(PgDadosCliente(Rst1.Fields("IdCliente")).doc, "###.###.###-##"), Format(PgDadosCliente(Rst1.Fields("IdCliente")).doc, "##.###.###/####-##"))
    rptPreVenda.Sections("Section2").Controls.Item("LblSRef").Caption = IIf(IsNull(Rst1.Fields("RefCliente")), "", Rst1.Fields("RefCliente"))
    '**************************************************************************************
    rptPreVenda.Sections("Section1").Controls.Item("txtQtd").DataField = "quantidade"
    rptPreVenda.Sections("Section1").Controls.Item("txtUnidade").DataField = "Unidade"
    rptPreVenda.Sections("Section1").Controls.Item("txtdescricao").DataField = "Descricao"
    rptPreVenda.Sections("Section1").Controls.Item("txtObs").DataField = "Obs"
    rptPreVenda.Sections("Section1").Controls.Item("txtNCM").DataField = "NCM"
    rptPreVenda.Sections("Section1").Controls.Item("txtICMS").DataField = "pICMS"
    rptPreVenda.Sections("Section1").Controls.Item("txtVlUnitario").DataField = "ValorUnitario"
    rptPreVenda.Sections("Section1").Controls.Item("txtipi").DataField = "ipi"
    'rptPreVenda.Sections("Section1").Controls.Item("txtTotalProduto").DataField = "VlProdBruto"
    'rptPreVenda.Sections("Section1").Controls.Item("txtTotalProduto").DataField = "SubTotal"
    rptPreVenda.Sections("Section1").Controls.Item("txtTotalProduto").DataField = "TotalProduto"
    '**************************************************************************************
    rptPreVenda.Sections("Section5").Controls.Item("lblFrete").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Frete")), "0,00", Rst1.Fields("Frete")))
    rptPreVenda.Sections("Section5").Controls.Item("lblSeguro").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Seguro")), "0,00", Rst1.Fields("Seguro")))
    rptPreVenda.Sections("Section5").Controls.Item("lblOutros").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Outros")), "0,00", Rst1.Fields("Outros")))
    rptPreVenda.Sections("Section5").Controls.Item("lblDesconto").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Desconto")), "0,00", Rst1.Fields("Desconto")))
    
    rptPreVenda.Sections("Section5").Controls.Item("lblvICMSST").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("vICMSST")), "0,00", Rst1.Fields("vICMSST")))
    
    rptPreVenda.Sections("Section5").Controls.Item("lblTotalPV").Caption = ConvMoeda(Rst1.Fields("VlTotalPV"))
    '**************************************************************************************
    rptPreVenda.Sections("Section5").Controls.Item("lblObs").Caption = IIf(IsNull(Rst1.Fields("Obs")), "", Rst1.Fields("Obs"))
    rptPreVenda.Sections("Section5").Controls.Item("lblPrazoEntrega").Caption = IIf(IsNull(Rst1.Fields("PrazoEntrega")), "", Rst1.Fields("PrazoEntrega"))
    rptPreVenda.Sections("Section5").Controls.Item("lblValidade").Caption = IIf(IsNull(Rst1.Fields("Validade")), "", Rst1.Fields("Validade"))
     
     
    If Rst1.Fields("transp_RetEnt") = 0 Then '0 - retira / 1- entrega
            '0 - Retira
            rptPreVenda.Sections("Section5").Controls.Item("lblTransp").Caption = Rst1.Fields("Cliente")
        Else
            '1 - Entrega
            If IsNull(Rst1.Fields("Transportadora")) Or Rst1.Fields("Transportadora") = 0 Then
                    rptPreVenda.Sections("Section5").Controls.Item("lblTransp").Caption = PgDadosEmpresa(ID_Empresa).Nome
                Else
                    rptPreVenda.Sections("Section5").Controls.Item("lblTransp").Caption = IIf(IsNull(Rst1.Fields("Transportadora")), " ", Rst1.Fields("Transportadora") & " - " & pgDadosTransportadora(Rst1.Fields("Transportadora")).Nome)
            End If
    End If
    
    
    rptPreVenda.Sections("Section5").Controls.Item("lblFreteConta").Caption = IIf(Rst1.Fields("FreteConta") = 0, "0 - Emitente", "1 - Destinatário")
    
    rptPreVenda.Sections("Section5").Controls.Item("lblVendedor").Caption = IIf(IsNull(Rst1.Fields("Vendedor")), "", PgDadosRhFuncionario(Rst1.Fields("Vendedor")).Assinatura)
    rptPreVenda.Sections("Section5").Controls.Item("lblCargo").Caption = IIf(IsNull(Rst1.Fields("Vendedor")), "", Trim(Mid(PgDadosRhFuncionario(Rst1.Fields("Vendedor")).Cargo, 5, Len(PgDadosRhFuncionario(Rst1.Fields("Vendedor")).Cargo))))
    rptPreVenda.Sections("Section5").Controls.Item("lblCondPagamento").Caption = IIf(IsNull(Rst1.Fields("CondicoesPagamento")), "", pgDescrCondPag(Rst1.Fields("CondicoesPagamento"))) & _
                                                                                  IIf(IsNull(Rst1.Fields("FormaPagamento")), "", " (" & pgDescrTipoDoc(Rst1.Fields("FormaPagamento")) & ")")
    'rptPreVenda.Sections("Section5").Controls.Item("fTotalPedido").DataField = "TotalProduto"
    
    rptPreVenda.Show 1
    '//rptPreVenda.PrintReport False, rptRangeAllPages
    

End Sub
Public Sub impPC(IdReg As Integer)
    'Pedido de Compra
    Dim sSQL        As String
    Dim Rst1        As Recordset
    Dim Rst2        As Recordset
    If IdReg = 0 Then Exit Sub
    sSQL = "SELECT * FROM estoquepedidocompra WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & IdReg
    Set Rst1 = RegistroBuscar(sSQL)
    'Set rptPedidoCompra.DataSource = Rst.DataSource
    
    sSQL = "SELECT * FROM estoquepedidocompraItens WHERE ID_Empresa = " & ID_Empresa & " AND IDPV = " & IdReg
    'sSQL = "SELECT FaturamentoPVItens.*, FaturamentoPVItens.SubTotal + FaturamentoPVItens.VlIPI AS VlProdBruto FROM FaturamentoPVItens WHERE IDPV = " & IdReg
    
    Set Rst2 = RegistroBuscar(sSQL)
    Set rptPedidoCompra.DataSource = Rst2.DataSource
    rptPedidoCompra.Title = "PV_" & Left(String(10, "0"), 10 - Len(IdReg)) & IdReg & "-" & Format(Rst1.Fields("Emissao"), "DD_MM_YYYY")
    DoEvents
    '**************************************************************************************
'    rptPedidoCompra.Sections("Section4").Controls.Item("LblTitulo").Font.Size = 12
'    rptPedidoCompra.Sections("Section4").Controls.Item("LblTitulo").Caption = "PEDIDO DE COMPRA"
    
    rptPedidoCompra.Sections("Section4").Controls.Item("LblEmissao").Caption = Rst1.Fields("Emissao") 'dtpEmissao.Value
    rptPedidoCompra.Sections("Section4").Controls.Item("LblNumero").Caption = Left(String(5, "0"), 5 - Len(IdReg)) & IdReg
    '**************************************************************************************
    rptPedidoCompra.Sections("Section2").Controls.Item("LblNome").Caption = Rst1.Fields("Cliente")
    rptPedidoCompra.Sections("Section2").Controls.Item("LblFone").Caption = IIf(IsNull(Rst1.Fields("Tel")), "", Rst1.Fields("Tel"))
    rptPedidoCompra.Sections("Section2").Controls.Item("LblEndCliente").Caption = PgDadosFornecedor(Rst1.Fields("IdCliente")).Lgr & " " & PgDadosFornecedor(Rst1.Fields("IdCliente")).Nro & " " & PgDadosFornecedor(Rst1.Fields("IdCliente")).Cpl
    rptPedidoCompra.Sections("Section2").Controls.Item("LblCNPJ").Caption = PgDadosFornecedor(Rst1.Fields("IdCliente")).Doc 'IIf(PgDadosFornecedor(Rst1.Fields("IdCliente")).pessoa = "Fisica", Format(PgDadosCliente(Rst1.Fields("IdCliente")).doc, "###.###.###-##"), Format(PgDadosCliente(Rst1.Fields("IdCliente")).doc, "##.###.###/####-##"))
    rptPedidoCompra.Sections("Section2").Controls.Item("LblSRef").Caption = IIf(IsNull(Rst1.Fields("RefCliente")), "", Rst1.Fields("RefCliente"))
    '**************************************************************************************
    rptPedidoCompra.Sections("Section1").Controls.Item("txtQtd").DataField = "quantidade"
    rptPedidoCompra.Sections("Section1").Controls.Item("txtUnidade").DataField = "Unidade"
    rptPedidoCompra.Sections("Section1").Controls.Item("txtdescricao").DataField = "Descricao"
    rptPedidoCompra.Sections("Section1").Controls.Item("txtObs").DataField = "Obs"
    rptPedidoCompra.Sections("Section1").Controls.Item("txtNCM").DataField = "NCM"
    rptPedidoCompra.Sections("Section1").Controls.Item("txtICMS").DataField = "pICMS"
    rptPedidoCompra.Sections("Section1").Controls.Item("txtVlUnitario").DataField = "ValorUnitario"
    rptPedidoCompra.Sections("Section1").Controls.Item("txtipi").DataField = "ipi"
    'rptPedidoCompra.Sections("Section1").Controls.Item("txtTotalProduto").DataField = "VlProdBruto"
    'rptPedidoCompra.Sections("Section1").Controls.Item("txtTotalProduto").DataField = "SubTotal"
    rptPedidoCompra.Sections("Section1").Controls.Item("txtTotalProduto").DataField = "TotalProduto"
    '**************************************************************************************
    rptPedidoCompra.Sections("Section5").Controls.Item("lblFrete").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Frete")), "0,00", Rst1.Fields("Frete")))
    rptPedidoCompra.Sections("Section5").Controls.Item("lblSeguro").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Seguro")), "0,00", Rst1.Fields("Seguro")))
    rptPedidoCompra.Sections("Section5").Controls.Item("lblOutros").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Outros")), "0,00", Rst1.Fields("Outros")))
    rptPedidoCompra.Sections("Section5").Controls.Item("lblDesconto").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Desconto")), "0,00", Rst1.Fields("Desconto")))
    
    rptPedidoCompra.Sections("Section5").Controls.Item("lblvICMSST").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("vICMSST")), "0,00", Rst1.Fields("vICMSST")))
    
    rptPedidoCompra.Sections("Section5").Controls.Item("lblTotalPV").Caption = ConvMoeda(Rst1.Fields("VlTotalPV"))
    '**************************************************************************************
    rptPedidoCompra.Sections("Section5").Controls.Item("lblObs").Caption = IIf(IsNull(Rst1.Fields("Obs")), "", Rst1.Fields("Obs"))
    rptPedidoCompra.Sections("Section5").Controls.Item("lblPrazoEntrega").Caption = IIf(IsNull(Rst1.Fields("PrazoEntrega")), "", Rst1.Fields("PrazoEntrega"))
    rptPedidoCompra.Sections("Section5").Controls.Item("lblValidade").Caption = IIf(IsNull(Rst1.Fields("Validade")), "", Rst1.Fields("Validade"))
     
     
    If Rst1.Fields("transp_RetEnt") = 0 Then '0 - retira / 1- entrega
            '0 - Retira
            rptPedidoCompra.Sections("Section5").Controls.Item("lblTransp").Caption = Rst1.Fields("Cliente")
        Else
            '1 - Entrega
            If IsNull(Rst1.Fields("Transportadora")) Or Rst1.Fields("Transportadora") = 0 Then
                    rptPedidoCompra.Sections("Section5").Controls.Item("lblTransp").Caption = PgDadosEmpresa(ID_Empresa).Nome
                Else
                    rptPedidoCompra.Sections("Section5").Controls.Item("lblTransp").Caption = IIf(IsNull(Rst1.Fields("Transportadora")), " ", Rst1.Fields("Transportadora") & " - " & pgDadosTransportadora(Rst1.Fields("Transportadora")).Nome)
            End If
    End If
    
    
    rptPedidoCompra.Sections("Section5").Controls.Item("lblFreteConta").Caption = IIf(Rst1.Fields("FreteConta") = 0, "0 - Emitente", "1 - Destinatário")
    
    rptPedidoCompra.Sections("Section5").Controls.Item("lblVendedor").Caption = IIf(IsNull(Rst1.Fields("Vendedor")), "", PgDadosRhFuncionario(Rst1.Fields("Vendedor")).Assinatura)
    rptPedidoCompra.Sections("Section5").Controls.Item("lblCargo").Caption = IIf(IsNull(Rst1.Fields("Vendedor")), "", Trim(Mid(PgDadosRhFuncionario(Rst1.Fields("Vendedor")).Cargo, 5, Len(PgDadosRhFuncionario(Rst1.Fields("Vendedor")).Cargo))))
    rptPedidoCompra.Sections("Section5").Controls.Item("lblCondPagamento").Caption = IIf(IsNull(Rst1.Fields("CondicoesPagamento")), "", pgDescrCondPag(Rst1.Fields("CondicoesPagamento"))) & _
                                                                                  IIf(IsNull(Rst1.Fields("FormaPagamento")), "", " (" & pgDescrTipoDoc(Rst1.Fields("FormaPagamento")) & ")")
    'rptPedidoCompra.Sections("Section5").Controls.Item("fTotalPedido").DataField = "TotalProduto"
    
    rptPedidoCompra.Show 1
    '//rptPedidoCompra.PrintReport False, rptRangeAllPages
    

End Sub
Public Sub impDuplicata(Id As Long, Optional Visualizar = True)
    'On Error Resume Next
    Dim Rst     As Recordset
    Dim extTMP  As String
    Dim sSQL    As String
    Dim nDup    As String

    sSQL = "SELECT * FROM FinanceiroContasPRCadastro WHERE ID_Empresa = " & ID_Empresa & " AND id = " & Id
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar registro da fatura", vbInformation, "Aviso"
            Exit Sub
        Else
            Rst.MoveFirst
    
    End If
    
    If PgDadosFinanceiroFatura(Id).ContaPR <> "R" Then
        MsgBox "Este tipo de documento so e permitido para documentos do tipo A RECEBER ou RECEBIDO!", vbInformation, "Aviso"
        Exit Sub
    End If
    
    nDup = PgDadosFinanceiroFatura(Id).NumDuplicata
    
    Set rptFatura.DataSource = Rst.DataSource
    rptFatura.Title = "Duplicata_" & nDup
    
    rptFatura.Sections("Section4").Controls.Item("Lb_Emissao").Caption = PgDadosFinanceiroFatura(Id).emissao
    rptFatura.Sections("Section4").Controls.Item("Lb_NFatura").Caption = PgDadosFinanceiroFatura(Id).NumFatura
    rptFatura.Sections("Section4").Controls.Item("Lb_VFatura").Caption = PgDadosFinanceiroFatura(Id).vlFatura
    rptFatura.Sections("Section4").Controls.Item("Lb_NDuplicata").Caption = nDup
    rptFatura.Sections("Section4").Controls.Item("Lb_VDuplicata").Caption = PgDadosFinanceiroFatura(Id).vlCobrado
    
    rptFatura.Sections("Section4").Controls.Item("Lb_Vencimento").Caption = PgDadosFinanceiroFatura(Id).Vencimento
    rptFatura.Sections("Section4").Controls.Item("Lb_Obs").Caption = PgDadosFinanceiroFatura(Id).Obs
    rptFatura.Sections("Section4").Controls.Item("Lb_Nome").Caption = PgDadosFinanceiroFatura(Id).Sacado
    rptFatura.Sections("Section4").Controls.Item("Lb_Endereco").Caption = PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Lgr & " " & PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Nro & " " & PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Cpl
    rptFatura.Sections("Section4").Controls.Item("Lb_Municipio").Caption = PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Mun
    rptFatura.Sections("Section4").Controls.Item("Lb_Estado").Caption = PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).uf
    rptFatura.Sections("Section4").Controls.Item("Lb_CEP").Caption = PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).CEP
    rptFatura.Sections("Section4").Controls.Item("Lb_CNPJ").Caption = PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Doc
    rptFatura.Sections("Section4").Controls.Item("Lb_IE").Caption = PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).IE
    
    extTMP = UCase(Extenso(Trim(ConvMoeda(PgDadosFinanceiroFatura(Id).vlCobrado)), "Reais", "Real")) & "  " & _
            Right(String(90, "*") & "  " & String(90, "*"), 180 - Len(UCase(Extenso(Trim(ConvMoeda(PgDadosFinanceiroFatura(Id).vlCobrado)), "Reais", "Real"))))
    
    rptFatura.Sections("Section4").Controls.Item("Lb_Extenso").Caption = extTMP
                 
                 
      '///////////// Segunda Duplicata \\\\\\\\\\\\\\\\\\\\\\
    'Set rptFatura.Sections("Section4").Controls.Item("imgLogo2").Picture = LoadPicture(PgDadosEmpresa(ID_Empresa).Logotipo)
    rptFatura.Sections("Section4").Controls.Item("lblCab11").Caption = rptFatura.Sections("Section4").Controls.Item("lblCab01").Caption
    rptFatura.Sections("Section4").Controls.Item("lblCab12").Caption = rptFatura.Sections("Section4").Controls.Item("lblCab02").Caption
    rptFatura.Sections("Section4").Controls.Item("lblCab13").Caption = rptFatura.Sections("Section4").Controls.Item("lblCab03").Caption
    rptFatura.Sections("Section4").Controls.Item("lblCab14").Caption = rptFatura.Sections("Section4").Controls.Item("lblCab04").Caption
    rptFatura.Sections("Section4").Controls.Item("Label37").Caption = rptFatura.Sections("Section4").Controls.Item("LB_Texto").Caption
    rptFatura.Sections("Section4").Controls.Item("Label30").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_Emissao").Caption
    rptFatura.Sections("Section4").Controls.Item("Label53").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_NFatura").Caption
    rptFatura.Sections("Section4").Controls.Item("Label54").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_VFatura").Caption
    rptFatura.Sections("Section4").Controls.Item("Label55").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_NDuplicata").Caption
    rptFatura.Sections("Section4").Controls.Item("Label56").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_VDuplicata").Caption
    rptFatura.Sections("Section4").Controls.Item("Label57").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_Obs").Caption
    rptFatura.Sections("Section4").Controls.Item("Label58").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_Vencimento").Caption
    rptFatura.Sections("Section4").Controls.Item("Label40").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_Nome").Caption
    rptFatura.Sections("Section4").Controls.Item("Label41").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_Endereco").Caption
    rptFatura.Sections("Section4").Controls.Item("Label42").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_Municipio").Caption
    rptFatura.Sections("Section4").Controls.Item("Label43").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_Estado").Caption
    rptFatura.Sections("Section4").Controls.Item("Label62").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_CEP").Caption
    rptFatura.Sections("Section4").Controls.Item("Label44").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_CNPJ").Caption
    rptFatura.Sections("Section4").Controls.Item("Label45").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_IE").Caption
    rptFatura.Sections("Section4").Controls.Item("Label46").Caption = rptFatura.Sections("Section4").Controls.Item("Lb_Extenso").Caption

  
    If Visualizar = True Then
            rptFatura.Show 1
        Else
            rptFatura.PrintReport True
    End If
    
'    rptFatura.Show 1

End Sub
Public Sub ImprBB_Pre(Id As Long)
    'Impressao de boleto Bancario pre-impresso
    
    If PgDadosFinanceiroFatura(Id).ContaPR <> "R" Then
        MsgBox "Este tipo de documento so e permitido para documentos do tipo A RECEBER ou RECEBIDO!", vbInformation, "Aviso"
        Exit Sub
    End If
    'formImpressoraSelecionar.SelecionarImpressora
     If formImpressoraSelecionar.SelecionarImpressora = False Then Exit Sub
    Printer.ScaleMode = 6
            Printer.FontName = "Tahoma"
            Printer.FontSize = 11
            Printer.FontItalic = False
            Printer.FontBold = False
            
        With Printer
            .CurrentX = 1
            .CurrentY = 18
            Printer.Print PgDadosFinanceiroFatura(Id).emissao
            
            .CurrentX = 135
            .CurrentY = 5
            Printer.Print PgDadosFinanceiroFatura(Id).Vencimento
            
            .CurrentX = 35
            .CurrentY = 18
            Printer.Print PgDadosFinanceiroFatura(Id).NumDuplicata
            
            .CurrentX = 135
            .CurrentY = 24
            Printer.Print PgDadosFinanceiroFatura(Id).vlDuplicata
            
            
            .CurrentX = 5
            .CurrentY = 33
            Printer.Print "Multa: " & PgDadosFinanceiroFatura(Id).Multa
            
            .CurrentX = 5
            .CurrentY = 38
            Printer.Print "Mora diaria: " & PgDadosFinanceiroFatura(Id).Juros & "% OU " & ConvMoeda(Val(PgDadosFinanceiroFatura(Id).Juros) * Val(ChkVal(PgDadosFinanceiroFatura(Id).vlDuplicata, 0, cDecMoeda)) / 100)
            '"Mora diaria: " & IIf(IsNull(Rst.Fields("cobr_Mora")), "", Rst.Fields("cobr_Mora")) & "% OU " & _

                        
            .CurrentX = 5
            .CurrentY = 43
            Printer.Print "Dias para Protesto: " & PgDadosFinanceiroFatura(Id).DiasProtesto
            
             
            .CurrentX = 5
            .CurrentY = 55
            Printer.Print PgDadosFinanceiroFatura(Id).CNPJSacado
            
            .CurrentX = 5
            .CurrentY = 63
            Printer.Print PgDadosFinanceiroFatura(Id).Sacado
            
            .CurrentX = 5
            .CurrentY = 68
            Printer.Print PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Lgr & " " & _
                          PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Nro & " " & _
                          PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Cpl & " - " & _
                          PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Bairro & " - " & _
                          PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Mun & "/" & _
                          PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).uf
                          
            
            Printer.EndDoc
        
        End With
End Sub
Public Sub ImprBB_Pre_Cont(sNFe As String)
    'Impressao de boleto Bancario pre-impresso Continuo
    'Uso somente na impressao direta da NFe (formFaturamentoNFeGerenciador
    On Error Resume Next
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim p       As Integer
    Dim idDupl  As Integer
    'Dim cDup    As Integer
    
    'formImpressoraSelecionar.SelecionarImpressora
    If formImpressoraSelecionar.SelecionarImpressora = False Then Exit Sub
    
    sSQL = "SELECT * FROM FaturamentoNFeCobranca WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & sNFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Exit Sub
    End If
    Rst.MoveFirst
    
    'cDup = 0
    p = 0
    Do Until Rst.EOF
        'cDup = cDup + 1
        Printer.ScaleMode = 6
        Printer.FontName = "Tahoma"
        Printer.FontSize = 11
        Printer.FontItalic = False
        Printer.FontBold = False
            
        With Printer
            .CurrentX = 1
            .CurrentY = 18 + p
            Printer.Print Rst.Fields("cobr_Emissao") ' PgDadosFinanceiroFatura(idDupl).Emissao
            
            .CurrentX = 135
            .CurrentY = 5 + p
            Printer.Print Rst.Fields("cobr_dVenc") 'PgDadosFinanceiroFatura(idDupl).Vencimento
            
            .CurrentX = 35
            .CurrentY = 18 + p
            Printer.Print Rst.Fields("cobr_nDup") 'PgDadosFinanceiroFatura(idDupl).NumDuplicata
            
            .CurrentX = 135
            .CurrentY = 24 + p
            Printer.Print Rst.Fields("cobr_vDup") 'PgDadosFinanceiroFatura(idDupl).vlDuplicata
            
            
            .CurrentX = 5
            .CurrentY = 33 + p
            Printer.Print "Multa: " & IIf(IsNull(Rst.Fields("cobr_Multa")), "0", Rst.Fields("cobr_Multa")) & "%" 'IIf(IsNull(Rst.Fields("cobr_Mora")), "", Rst.Fields("cobr_Mora")) & "% OU " & _
                            ConvMoeda(Val(IIf(IsNull(Rst.Fields("cobr_Mora")), 0, Rst.Fields("cobr_Mora"))) * Val(ChkVal(Rst.Fields("cobr_vDup"), 0, cDecMoeda)) / 100)
            
            .CurrentX = 5
            .CurrentY = 38 + p
            Printer.Print "Mora diaria: " & IIf(IsNull(Rst.Fields("cobr_Mora")), "", Rst.Fields("cobr_Mora")) & "% OU " & _
                            ConvMoeda(Val(IIf(IsNull(Rst.Fields("cobr_Mora")), 0, Rst.Fields("cobr_Mora"))) * Val(ChkVal(Rst.Fields("cobr_vDup"), 0, cDecMoeda)) / 100)
                        
            .CurrentX = 5
            .CurrentY = 43 + p
            Printer.Print "Dias para Protesto: " & Rst.Fields("cobr_protesto") 'pgDadosConta(PgDadosFinanceiroFatura(idDupl).IdConta).DiasProtesto
            
            '.CurrentX = 5
            '.CurrentY = 43 + p
            'Printer.Print PgDadosFinanceiroFatura(idDupl).Obs
             
            .CurrentX = 5
            .CurrentY = 55 + p
            Printer.Print PgDadosCliente(Rst.Fields("cobr_idCliente")).Doc 'PgDadosFinanceiroFatura(idDupl).CNPJSacado
            
            .CurrentX = 5
            .CurrentY = 63 + p
            Printer.Print PgDadosCliente(Rst.Fields("cobr_idCliente")).Nome 'PgDadosFinanceiroFatura(idDupl).Sacado
            
            .CurrentX = 5
            .CurrentY = 68 + p
            Printer.Print PgDadosCliente(Rst.Fields("cobr_idCliente")).Lgr & " " & _
                          PgDadosCliente(Rst.Fields("cobr_idCliente")).Nro & " " & _
                          PgDadosCliente(Rst.Fields("cobr_idCliente")).Cpl & " - " & _
                          PgDadosCliente(Rst.Fields("cobr_idCliente")).Bairro & " - " & _
                          PgDadosCliente(Rst.Fields("cobr_idCliente")).Mun & "/" & _
                          PgDadosCliente(Rst.Fields("cobr_idCliente")).uf
            Rst.MoveNext
            p = p + 102
         End With
         'If cDup = 3 Then
         '   cDup = 0
         '   p = 0
         '   'Printer.EndDoc
         '   Printer.NewPage
        'End If
    Loop
    Printer.EndDoc
        
       
End Sub

Public Sub ImprBoletoBancario(Id As Long, Optional Visualizar = True) ', NossoNumero As String, LinhaDigitavel As String, CodigoBarras As String)
    Dim Rst             As Recordset
    Dim sSQL            As String
    
    
    
    sSQL = "SELECT * FROM FinanceiroContasPRCadastro WHERE ID_Empresa = " & ID_Empresa & " AND id =" & Id
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar o Boleto para impressão.", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            '#################################################################################################################
            '### Dados do Boleto
            Dim NossoNum        As String
            Dim AgenciaConta    As String
            Dim sMulta          As String
            Dim MoraDiaria      As String
            Dim nDup            As String
            
            
            
            NossoNum = Trim(PgDadosFinanceiroFatura(Id).NossoNumero)
            'NossoNum = Mid(NossoNum, 1, Len(NossoNum) - 1) & "-" & Right(NossoNum, 1)
            
            AgenciaConta = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).agencia & IIf(pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).AgenciaDV <> "", "-" & pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).AgenciaDV, "") & " / " & pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).conta & IIf(pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).ContaDV <> "", "-" & pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).ContaDV, "")
            
            'Mora Diaria
            MoraDiaria = ConvMoeda(cobCalcMora(PgDadosFinanceiroFatura(Id).vlDuplicata, 1, PgDadosFinanceiroFatura(Id).Juros, "D"))
            sMulta = PgDadosFinanceiroFatura(Id).Multa & "% ou " & ConvMoeda(cobCalcMulta(PgDadosFinanceiroFatura(Id).vlDuplicata, PgDadosFinanceiroFatura(Id).Multa, 1))
            
            nDup = PgDadosFinanceiroFatura(Id).NumDuplicata
            
            '#################################################################################################################
            Set rptBoletoBancario.DataSource = Rst.DataSource
            
            rptBoletoBancario.Title = "Boleto_" & nDup
            
            With rptBoletoBancario.Sections("Section1").Controls
                '************************* RECIBO CEDENTE *******************************************************************
                .Item("lblBCO1").Caption = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Nome
                .Item("lblBCOc1").Caption = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero & "-" & Trim(calculo_dv11base9(pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero))
                .Item("lblCedente1").Caption = PgDadosEmpresa(ID_Empresa).Nome
                
                .Item("lblDE1").Caption = PgDadosFinanceiroFatura(Id).emissao
                .Item("lblV1").Caption = PgDadosFinanceiroFatura(Id).Vencimento
                .Item("lblAC1").Caption = AgenciaConta
                .Item("lblND1").Caption = nDup
                .Item("lblDTP1").Caption = Date
                .Item("lblNN1").Caption = NossoNum
                '.Item("lblVD1").Caption = PgDadosFinanceiroFatura(Id).vlDuplicata
                .Item("lblVD1").Caption = ChkVal(IIf(ChkVal(PgDadosFinanceiroFatura(Id).vlCobrado, 0, cDecMoeda) = 0, PgDadosFinanceiroFatura(Id).vlDuplicata, PgDadosFinanceiroFatura(Id).vlCobrado), 0, cDecMoeda)
                .Item("lblCA1").Caption = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).carteira & " " & pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).Variacao
                .Item("lblLD1").Caption = Formatar_Linha_Digitavel(PgDadosFinanceiroFatura(Id).LinhaDigitavel)

                .Item("lblNome1").Caption = PgDadosFinanceiroFatura(Id).Sacado
                .Item("lblDoc1").Caption = PgDadosFinanceiroFatura(Id).CNPJSacado
                
                 
                .Item("lblMsg11").Caption = "Multa: " & sMulta
                .Item("lblMsg12").Caption = "Mora diaria: " & PgDadosFinanceiroFatura(Id).Juros & "% OU " & MoraDiaria
                'multa
                .Item("lblMsg13").Caption = IIf(Trim(PgDadosFinanceiroFatura(Id).DiasProtesto) = "0", "", "Dias para Protesto: " & PgDadosFinanceiroFatura(Id).DiasProtesto)
                .Item("lblMsg14").Caption = PgDadosFinanceiroFatura(Id).ObsBol1
                .Item("lblMsg15").Caption = PgDadosFinanceiroFatura(Id).ObsBol2
                .Item("lblMsg16").Caption = PgDadosFinanceiroFatura(Id).ObsBol3
                '************************************************************************************************************
            'End With
            'With rptBoletoBancario.Sections("Boleto2").Controls
                '************************* RECIBO SACADO *******************************************************************
                .Item("lblLD2").Caption = Formatar_Linha_Digitavel(PgDadosFinanceiroFatura(Id).LinhaDigitavel)
                .Item("lblBCO2").Caption = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Nome
                .Item("lblBCOc2").Caption = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero & "-" & Trim(calculo_dv11base9(pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero))
                .Item("lblCedente2").Caption = PgDadosEmpresa(ID_Empresa).Nome
                
                .Item("lblDE2").Caption = PgDadosFinanceiroFatura(Id).emissao
                .Item("lblV2").Caption = PgDadosFinanceiroFatura(Id).Vencimento
                .Item("lblAC2").Caption = AgenciaConta
                .Item("lblND2").Caption = nDup
                .Item("lblDTP2").Caption = Date
                .Item("lblNN2").Caption = NossoNum
                .Item("lblCA2").Caption = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).carteira & " " & pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).Variacao
                .Item("lblVD2").Caption = PgDadosFinanceiroFatura(Id).vlDuplicata
                
                .Item("lblDA2").Caption = IIf(ChkVal(PgDadosFinanceiroFatura(Id).Abatimento, 0, cDecMoeda) = 0, "", ChkVal(PgDadosFinanceiroFatura(Id).Abatimento, 0, cDecMoeda))
                .Item("lblOD2").Caption = IIf(ChkVal(PgDadosFinanceiroFatura(Id).Deducoes, 0, cDecMoeda) = 0, "", ChkVal(PgDadosFinanceiroFatura(Id).Deducoes, 0, cDecMoeda))
                .Item("lblJM2").Caption = IIf(ChkVal(PgDadosFinanceiroFatura(Id).MultaMora, 0, cDecMoeda) = 0, "", ChkVal(PgDadosFinanceiroFatura(Id).MultaMora, 0, cDecMoeda))
                .Item("lblOA2").Caption = IIf(ChkVal(PgDadosFinanceiroFatura(Id).Acrescimo, 0, cDecMoeda) = 0, "", ChkVal(PgDadosFinanceiroFatura(Id).Acrescimo, 0, cDecMoeda))
                .Item("lblVC2").Caption = IIf(ChkVal(PgDadosFinanceiroFatura(Id).vlCobrado, 0, cDecMoeda) = 0, "", ChkVal(PgDadosFinanceiroFatura(Id).vlCobrado, 0, cDecMoeda))
                
                .Item("lblMsg21").Caption = "Multa: " & sMulta 'PgDadosFinanceiroFatura(Id).Multa
                .Item("lblMsg22").Caption = "Mora diaria: " & PgDadosFinanceiroFatura(Id).Juros & "% OU " & MoraDiaria
                'Multa
                .Item("lblMsg23").Caption = IIf(Trim(PgDadosFinanceiroFatura(Id).DiasProtesto) = "0", "", "Dias para Protesto: " & PgDadosFinanceiroFatura(Id).DiasProtesto)
                .Item("lblMsg24").Caption = PgDadosFinanceiroFatura(Id).ObsBol1
                .Item("lblMsg25").Caption = PgDadosFinanceiroFatura(Id).ObsBol2
                .Item("lblMsg26").Caption = PgDadosFinanceiroFatura(Id).ObsBol3
                
                .Item("lblNome2").Caption = PgDadosFinanceiroFatura(Id).Sacado
                .Item("lblDoc2").Caption = PgDadosFinanceiroFatura(Id).CNPJSacado
                '.Item("lblCB1").Caption = CodigoBarras
                '************************************************************************************************************

                '************************* FICHA COMPENSACAO *******************************************************************
                .Item("lblLD3").Caption = Formatar_Linha_Digitavel(PgDadosFinanceiroFatura(Id).LinhaDigitavel)
                .Item("lblBCO3").Caption = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Nome
                .Item("lblBCOc3").Caption = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero & "-" & Trim(calculo_dv11base9(pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero))
                .Item("lblCedente3").Caption = PgDadosEmpresa(ID_Empresa).Nome
                .Item("lblDE3").Caption = PgDadosFinanceiroFatura(Id).emissao
                .Item("lblV3").Caption = PgDadosFinanceiroFatura(Id).Vencimento
                .Item("lblAC3").Caption = AgenciaConta
                .Item("lblND3").Caption = nDup
                .Item("lblDTP3").Caption = Date
                .Item("lblNN3").Caption = NossoNum
                .Item("lblCA3").Caption = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).carteira & " " & pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).Variacao
                .Item("lblVD3").Caption = PgDadosFinanceiroFatura(Id).vlDuplicata
                
                .Item("lblDA3").Caption = IIf(ChkVal(PgDadosFinanceiroFatura(Id).Abatimento, 0, cDecMoeda) = 0, "", ChkVal(PgDadosFinanceiroFatura(Id).Abatimento, 0, cDecMoeda))
                .Item("lblOD3").Caption = IIf(ChkVal(PgDadosFinanceiroFatura(Id).Deducoes, 0, cDecMoeda) = 0, "", ChkVal(PgDadosFinanceiroFatura(Id).Deducoes, 0, cDecMoeda))
                .Item("lblJM3").Caption = IIf(ChkVal(PgDadosFinanceiroFatura(Id).MultaMora, 0, cDecMoeda) = 0, "", ChkVal(PgDadosFinanceiroFatura(Id).MultaMora, 0, cDecMoeda))
                .Item("lblOA3").Caption = IIf(ChkVal(PgDadosFinanceiroFatura(Id).Acrescimo, 0, cDecMoeda) = 0, "", ChkVal(PgDadosFinanceiroFatura(Id).Acrescimo, 0, cDecMoeda))
                .Item("lblVC3").Caption = IIf(ChkVal(PgDadosFinanceiroFatura(Id).vlCobrado, 0, cDecMoeda) = 0, "", ChkVal(PgDadosFinanceiroFatura(Id).vlCobrado, 0, cDecMoeda))
                
                
                .Item("lblMsg31").Caption = "Multa: " & sMulta 'PgDadosFinanceiroFatura(Id).Multa
                .Item("lblMsg32").Caption = "Mora diaria: " & PgDadosFinanceiroFatura(Id).Juros & "% OU " & MoraDiaria
                .Item("lblMsg33").Caption = IIf(Trim(PgDadosFinanceiroFatura(Id).DiasProtesto) = "0", "", "Dias para Protesto: " & PgDadosFinanceiroFatura(Id).DiasProtesto)
                .Item("lblMsg34").Caption = PgDadosFinanceiroFatura(Id).ObsBol1
                .Item("lblMsg35").Caption = PgDadosFinanceiroFatura(Id).ObsBol2
                .Item("lblMsg36").Caption = PgDadosFinanceiroFatura(Id).ObsBol3
                
                .Item("lblNome3").Caption = PgDadosFinanceiroFatura(Id).Sacado
                .Item("lblDoc3").Caption = PgDadosFinanceiroFatura(Id).CNPJSacado
                .Item("lblCB3").Caption = PgDadosFinanceiroFatura(Id).CodigoBarras
                '************************************************************************************************************
            End With
            If Visualizar = True Then
                    rptBoletoBancario.Show 1
                Else
                    rptBoletoBancario.PrintReport False
            End If
    End If
    Rst.Close
End Sub
Private Function Formatar_Linha_Digitavel(sequencia As String) As String

    Dim seq1        As String
    Dim seq2        As String
    Dim seq3        As String
    Dim seq4        As String
    Dim seq5        As String
    
    seq1 = Mid(sequencia, 1, 10)
    seq1 = Left(seq1, 5) & "." & Right(seq1, 5)
    
    seq2 = Mid(sequencia, 11, 11)
    seq2 = Left(seq2, 5) & "." & Right(seq2, 6)

    seq3 = Mid(sequencia, 22, 11)
    seq3 = Left(seq3, 5) & "." & Right(seq3, 6)
    
    seq4 = Mid(sequencia, 33, 1)
    
    seq5 = Mid(sequencia, 34, Len(sequencia))
    
    
    Formatar_Linha_Digitavel = seq1 & " " & seq2 & " " & seq3 & " " & seq4 & " " & seq5

End Function

Public Sub ImpRomaneio(IdReg As Integer)
    Dim sSQL        As String
    Dim Rst1        As Recordset
    Dim Rst2        As Recordset
    If IdReg = 0 Then Exit Sub
    sSQL = "SELECT * FROM FaturamentoPV WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & IdReg
    Set Rst1 = RegistroBuscar(sSQL)
    'Set rptRomaneio.DataSource = Rst.DataSource
    
    sSQL = "SELECT * FROM FaturamentoPVItens WHERE ID_Empresa = " & ID_Empresa & " AND IDPV = " & IdReg
    'sSQL = "SELECT FaturamentoPVItens.*, FaturamentoPVItens.SubTotal + FaturamentoPVItens.VlIPI AS VlProdBruto FROM FaturamentoPVItens WHERE IDPV = " & IdReg
    
    Set Rst2 = RegistroBuscar(sSQL)
    Set rptRomaneio.DataSource = Rst2.DataSource
    rptRomaneio.Title = "PV_" & Left(String(10, "0"), 10 - Len(IdReg)) & IdReg & "-" & Format(Rst1.Fields("Emissao"), "DD_MM_YYYY")
    DoEvents
    '**************************************************************************************
'    rptRomaneio.Sections("Section4").Controls.Item("LblTitulo").Caption = "PRE-VENDA"
    rptRomaneio.Sections("Section4").Controls.Item("LblEmissao").Caption = Rst1.Fields("Emissao") 'dtpEmissao.Value
    rptRomaneio.Sections("Section4").Controls.Item("LblNumero").Caption = Left(String(5, "0"), 5 - Len(IdReg)) & IdReg
    '**************************************************************************************
    rptRomaneio.Sections("Section2").Controls.Item("LblNome").Caption = Rst1.Fields("Cliente")
    rptRomaneio.Sections("Section2").Controls.Item("LblFone").Caption = IIf(IsNull(Rst1.Fields("Tel")), "", Rst1.Fields("Tel"))
    rptRomaneio.Sections("Section2").Controls.Item("LblEndCliente").Caption = PgDadosCliente(Rst1.Fields("IdCliente")).Lgr & " " & PgDadosCliente(Rst1.Fields("IdCliente")).Nro & " " & PgDadosCliente(Rst1.Fields("IdCliente")).Cpl
    rptRomaneio.Sections("Section2").Controls.Item("LblCNPJ").Caption = PgDadosCliente(Rst1.Fields("IdCliente")).Doc 'IIf(PgDadosCliente(Rst1.Fields("IdCliente")).pessoa = "Fisica", Format(PgDadosCliente(Rst1.Fields("IdCliente")).doc, "###.###.###-##"), Format(PgDadosCliente(Rst1.Fields("IdCliente")).doc, "##.###.###/####-##"))
    rptRomaneio.Sections("Section2").Controls.Item("LblSRef").Caption = IIf(IsNull(Rst1.Fields("RefCliente")), "", Rst1.Fields("RefCliente"))
    '**************************************************************************************
    rptRomaneio.Sections("Section1").Controls.Item("txtQtd").DataField = "quantidade"
    rptRomaneio.Sections("Section1").Controls.Item("txtUnidade").DataField = "Unidade"
    rptRomaneio.Sections("Section1").Controls.Item("txtdescricao").DataField = "Descricao"
    rptRomaneio.Sections("Section1").Controls.Item("txtObs").DataField = "Obs"
    rptRomaneio.Sections("Section1").Controls.Item("txtNCM").DataField = "NCM"
'    rptRomaneio.Sections("Section1").Controls.Item("txtICMS").DataField = "pICMS"
'    rptRomaneio.Sections("Section1").Controls.Item("txtVlUnitario").DataField = "ValorUnitario"
'    rptRomaneio.Sections("Section1").Controls.Item("txtipi").DataField = "ipi"
'    rptRomaneio.Sections("Section1").Controls.Item("txtTotalProduto").DataField = "VlProdBruto"
'    rptRomaneio.Sections("Section1").Controls.Item("txtTotalProduto").DataField = "SubTotal"
'    rptRomaneio.Sections("Section1").Controls.Item("txtTotalProduto").DataField = "TotalProduto"
    '**************************************************************************************
'    rptRomaneio.Sections("Section5").Controls.Item("lblFrete").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Frete")), "0,00", Rst1.Fields("Frete")))
'    rptRomaneio.Sections("Section5").Controls.Item("lblSeguro").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Seguro")), "0,00", Rst1.Fields("Seguro")))
'    rptRomaneio.Sections("Section5").Controls.Item("lblOutros").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Outros")), "0,00", Rst1.Fields("Outros")))
'    rptRomaneio.Sections("Section5").Controls.Item("lblDesconto").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Desconto")), "0,00", Rst1.Fields("Desconto")))
'
'    rptRomaneio.Sections("Section5").Controls.Item("lblvICMSST").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("vICMSST")), "0,00", Rst1.Fields("vICMSST")))
'
'    rptRomaneio.Sections("Section5").Controls.Item("lblTotalPV").Caption = ConvMoeda(Rst1.Fields("VlTotalPV"))
    '**************************************************************************************
    rptRomaneio.Sections("Section5").Controls.Item("lblObs").Caption = IIf(IsNull(Rst1.Fields("Obs")), "", Rst1.Fields("Obs"))
    rptRomaneio.Sections("Section5").Controls.Item("lblPrazoEntrega").Caption = IIf(IsNull(Rst1.Fields("PrazoEntrega")), "", Rst1.Fields("PrazoEntrega"))
    rptRomaneio.Sections("Section5").Controls.Item("lblValidade").Caption = IIf(IsNull(Rst1.Fields("Validade")), "", Rst1.Fields("Validade"))
     
     
    If Rst1.Fields("transp_RetEnt") = 0 Then '0 - retira / 1- entrega
            '0 - Retira
            rptRomaneio.Sections("Section5").Controls.Item("lblTransp").Caption = Rst1.Fields("Cliente")
        Else
            '1 - Entrega
            If IsNull(Rst1.Fields("Transportadora")) Or Rst1.Fields("Transportadora") = 0 Then
                    rptRomaneio.Sections("Section5").Controls.Item("lblTransp").Caption = PgDadosEmpresa(ID_Empresa).Nome
                Else
                    rptRomaneio.Sections("Section5").Controls.Item("lblTransp").Caption = IIf(IsNull(Rst1.Fields("Transportadora")), " ", Rst1.Fields("Transportadora") & " - " & pgDadosTransportadora(Rst1.Fields("Transportadora")).Nome)
            End If
    End If
    
    
    rptRomaneio.Sections("Section5").Controls.Item("lblFreteConta").Caption = IIf(Rst1.Fields("FreteConta") = 0, "0 - Emitente", "1 - Destinatário")
    
    rptRomaneio.Sections("Section5").Controls.Item("lblVendedor").Caption = IIf(IsNull(Rst1.Fields("Vendedor")), "", PgDadosRhFuncionario(Rst1.Fields("Vendedor")).Assinatura)
    rptRomaneio.Sections("Section5").Controls.Item("lblCargo").Caption = IIf(IsNull(Rst1.Fields("Vendedor")), "", Trim(Mid(PgDadosRhFuncionario(Rst1.Fields("Vendedor")).Cargo, 5, Len(PgDadosRhFuncionario(Rst1.Fields("Vendedor")).Cargo))))
    rptRomaneio.Sections("Section5").Controls.Item("lblCondPagamento").Caption = ""
'    rptRomaneio.Sections("Section5").Controls.Item("lblCondPagamento").Caption = IIf(IsNull(Rst1.Fields("CondicoesPagamento")), "", pgDescrCondPag(Rst1.Fields("CondicoesPagamento"))) & _
                                                                                  IIf(IsNull(Rst1.Fields("FormaPagamento")), "", " (" & pgDescrTipoDoc(Rst1.Fields("FormaPagamento")) & ")")
'    rptRomaneio.Sections("Section5").Controls.Item("fTotalPedido").DataField = "TotalProduto"
    
    rptRomaneio.Show 1
    '//rptRomaneio.PrintReport False, rptRangeAllPages
    

End Sub


Public Sub ImprimirListaClientes(Optional numVendedor As Integer, Optional uf As String)
    'Listagem de Cientes
    Dim sSQL        As String
    Dim Rst        As Recordset
    'Dim Rst2        As Recordset
    'If idReg = 0 Then Exit Sub
    
    sSQL = "SELECT * FROM clientes WHERE ID_Empresa = " & ID_Empresa & _
    IIf(numVendedor = 0, "", " AND vendedor = '" & numVendedor & "'") & _
    IIf(Len(uf) = 0, "", " AND UF = '" & uf & "'")
    '& " AND ID = " & idReg
    Set Rst = RegistroBuscar(sSQL)
    'Set rptClienteListagem.DataSource = Rst.DataSource
    
    'sSQL = "SELECT * FROM FaturamentoPVItens WHERE ID_Empresa = " & ID_Empresa & " AND IDPV = " & idReg
    'sSQL = "SELECT FaturamentoPVItens.*, FaturamentoPVItens.SubTotal + FaturamentoPVItens.VlIPI AS VlProdBruto FROM FaturamentoPVItens WHERE IDPV = " & IdReg
    
    'Set Rst2 = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        MsgBox "Nenhum registro encontrado!", vbInformation, App.EXEName
        Rst.Close
        Exit Sub
    End If
    Set rptClienteListagem.DataSource = Rst.DataSource
  
    
    
    rptClienteListagem.Title = "Listagem de Clientes " ' & Left(String(10, "0"), 10 - Len(idReg)) & idReg & "-" & Format(Rst1.Fields("Emissao"), "DD_MM_YYYY")
    DoEvents
'    '**************************************************************************************
'    rptClienteListagem.Sections("Section4").Controls.Item("LblTitulo").Caption = tpDoc
'
'    rptClienteListagem.Sections("Section4").Controls.Item("LblEmissao").Caption = Rst1.Fields("Emissao") 'dtpEmissao.Value
'    rptClienteListagem.Sections("Section4").Controls.Item("LblNumero").Caption = Left(String(5, "0"), 5 - Len(idReg)) & idReg
'    '**************************************************************************************
'    rptClienteListagem.Sections("Section2").Controls.Item("LblNome").Caption = Rst1.Fields("Cliente")
'    rptClienteListagem.Sections("Section2").Controls.Item("LblFone").Caption = IIf(IsNull(Rst1.Fields("Tel")), "", Rst1.Fields("Tel"))
'    rptClienteListagem.Sections("Section2").Controls.Item("LblEndCliente").Caption = PgDadosCliente(Rst1.Fields("IdCliente")).Lgr & " " & PgDadosCliente(Rst1.Fields("IdCliente")).Nro & " " & PgDadosCliente(Rst1.Fields("IdCliente")).Cpl
'    rptClienteListagem.Sections("Section2").Controls.Item("LblCNPJ").Caption = PgDadosCliente(Rst1.Fields("IdCliente")).Doc 'IIf(PgDadosCliente(Rst1.Fields("IdCliente")).pessoa = "Fisica", Format(PgDadosCliente(Rst1.Fields("IdCliente")).doc, "###.###.###-##"), Format(PgDadosCliente(Rst1.Fields("IdCliente")).doc, "##.###.###/####-##"))
'    rptClienteListagem.Sections("Section2").Controls.Item("LblSRef").Caption = IIf(IsNull(Rst1.Fields("RefCliente")), "", Rst1.Fields("RefCliente"))
'    '**************************************************************************************
'    rptClienteListagem.Sections("Section1").Controls.Item("txtQtd").DataField = "quantidade"
'    rptClienteListagem.Sections("Section1").Controls.Item("txtUnidade").DataField = "Unidade"
'    rptClienteListagem.Sections("Section1").Controls.Item("txtdescricao").DataField = "Descricao"
'    rptClienteListagem.Sections("Section1").Controls.Item("txtObs").DataField = "Obs"
'    rptClienteListagem.Sections("Section1").Controls.Item("txtNCM").DataField = "NCM"
'    rptClienteListagem.Sections("Section1").Controls.Item("txtICMS").DataField = "pICMS"
'    rptClienteListagem.Sections("Section1").Controls.Item("txtVlUnitario").DataField = "ValorUnitario"
'    rptClienteListagem.Sections("Section1").Controls.Item("txtipi").DataField = "ipi"
'    'rptClienteListagem.Sections("Section1").Controls.Item("txtTotalProduto").DataField = "VlProdBruto"
'    'rptClienteListagem.Sections("Section1").Controls.Item("txtTotalProduto").DataField = "SubTotal"
'    rptClienteListagem.Sections("Section1").Controls.Item("txtTotalProduto").DataField = "TotalProduto"
'    '**************************************************************************************
'    rptClienteListagem.Sections("Section5").Controls.Item("lblFrete").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Frete")), "0,00", Rst1.Fields("Frete")))
'    rptClienteListagem.Sections("Section5").Controls.Item("lblSeguro").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Seguro")), "0,00", Rst1.Fields("Seguro")))
'    rptClienteListagem.Sections("Section5").Controls.Item("lblOutros").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Outros")), "0,00", Rst1.Fields("Outros")))
'    rptClienteListagem.Sections("Section5").Controls.Item("lblDesconto").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("Desconto")), "0,00", Rst1.Fields("Desconto")))
'
'    rptClienteListagem.Sections("Section5").Controls.Item("lblvICMSST").Caption = ConvMoeda(IIf(IsNull(Rst1.Fields("vICMSST")), "0,00", Rst1.Fields("vICMSST")))
'
'    rptClienteListagem.Sections("Section5").Controls.Item("lblTotalPV").Caption = ConvMoeda(Rst1.Fields("VlTotalPV"))
'    '**************************************************************************************
'    rptClienteListagem.Sections("Section5").Controls.Item("lblObs").Caption = IIf(IsNull(Rst1.Fields("Obs")), "", Rst1.Fields("Obs"))
'    rptClienteListagem.Sections("Section5").Controls.Item("lblPrazoEntrega").Caption = IIf(IsNull(Rst1.Fields("PrazoEntrega")), "", Rst1.Fields("PrazoEntrega"))
'    rptClienteListagem.Sections("Section5").Controls.Item("lblValidade").Caption = IIf(IsNull(Rst1.Fields("Validade")), "", Rst1.Fields("Validade"))
    rptClienteListagem.Show 1
     
    
End Sub
