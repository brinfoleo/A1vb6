Attribute VB_Name = "ModuloBoleto"
Option Explicit
'RJ, 13.09.2017
'Modulo que usa a DLL feita em c# para emissao
'do cnab240 e boletos



Public Sub DllMontarBoleto(idFatura As Integer)
'On Error GoTo trterro
'
'   Rst.MoveFirst
'            '#################################################################################################################
'            '### Dados do Boleto
'            Dim NossoNum        As String
'            Dim AgenciaConta    As String
'            Dim sMulta          As String
'            Dim MoraDiaria      As String
'            Dim nDup            As String
'
'
'
'            NossoNum = Trim(PgDadosFinanceiroFatura(id).NossoNumero)
'            NossoNum = Mid(NossoNum, 1, Len(NossoNum) - 1) & "-" & Right(NossoNum, 1)
'
'            AgenciaConta = pgDadosConta(PgDadosFinanceiroFatura(id).idConta).agencia & IIf(pgDadosConta(PgDadosFinanceiroFatura(id).idConta).AgenciaDV <> "", "-" & pgDadosConta(PgDadosFinanceiroFatura(id).idConta).AgenciaDV, "") & " / " & pgDadosConta(PgDadosFinanceiroFatura(id).idConta).conta & IIf(pgDadosConta(PgDadosFinanceiroFatura(id).idConta).ContaDV <> "", "-" & pgDadosConta(PgDadosFinanceiroFatura(id).idConta).ContaDV, "")
'
'            'Mora Diaria
'            MoraDiaria = ConvMoeda(cobCalcMora(PgDadosFinanceiroFatura(id).vlDuplicata, 1, PgDadosFinanceiroFatura(id).Juros, "D"))
'            sMulta = PgDadosFinanceiroFatura(id).Multa & "% ou " & ConvMoeda(cobCalcMulta(PgDadosFinanceiroFatura(id).vlDuplicata, PgDadosFinanceiroFatura(id).Multa, 1))
'
'            nDup = PgDadosFinanceiroFatura(id).NumDuplicata
'
'            '#################################################################################################################
'            Set rptBoletoBancario.DataSource = Rst.DataSource
'
'            rptBoletoBancario.Title = "Boleto_" & nDup
'
'            With rptBoletoBancario.Sections("Section1").Controls
'                '************************* RECIBO CEDENTE *******************************************************************
'                .Item("lblBCO1").Caption = pgDadosBanco(PgDadosFinanceiroFatura(id).IdBanco).Nome
'                .Item("lblBCOc1").Caption = pgDadosBanco(PgDadosFinanceiroFatura(id).IdBanco).Numero & "-" & Trim(calculo_dv11base9(pgDadosBanco(PgDadosFinanceiroFatura(id).IdBanco).Numero))
'                .Item("lblCedente1").Caption = PgDadosEmpresa(ID_Empresa).Nome
'
'                .Item("lblDE1").Caption = PgDadosFinanceiroFatura(id).emissao
'                .Item("lblV1").Caption = PgDadosFinanceiroFatura(id).Vencimento
'                .Item("lblAC1").Caption = AgenciaConta
'                .Item("lblND1").Caption = nDup
'                .Item("lblDTP1").Caption = Date
'                .Item("lblNN1").Caption = NossoNum
'                '.Item("lblVD1").Caption = PgDadosFinanceiroFatura(Id).vlDuplicata
'                .Item("lblVD1").Caption = ChkVal(IIf(ChkVal(PgDadosFinanceiroFatura(id).vlCobrado, 0, cDecMoeda) = 0, PgDadosFinanceiroFatura(id).vlDuplicata, PgDadosFinanceiroFatura(id).vlCobrado), 0, cDecMoeda)
'                .Item("lblCA1").Caption = pgDadosConta(PgDadosFinanceiroFatura(id).idConta).carteira & " " & pgDadosConta(PgDadosFinanceiroFatura(id).idConta).Variacao
'                .Item("lblLD1").Caption = Formatar_Linha_Digitavel(PgDadosFinanceiroFatura(id).LinhaDigitavel)
'
'                .Item("lblNome1").Caption = PgDadosFinanceiroFatura(id).Sacado
'                .Item("lblDoc1").Caption = PgDadosFinanceiroFatura(id).CNPJSacado
'
'
'                .Item("lblMsg11").Caption = "Multa: " & sMulta
'                .Item("lblMsg12").Caption = "Mora diaria: " & PgDadosFinanceiroFatura(id).Juros & "% OU " & MoraDiaria
'                'multa
'                .Item("lblMsg13").Caption = IIf(Trim(PgDadosFinanceiroFatura(id).DiasProtesto) = "0", "", "Dias para Protesto: " & PgDadosFinanceiroFatura(id).DiasProtesto)
'                .Item("lblMsg14").Caption = PgDadosFinanceiroFatura(id).ObsBol1
'                .Item("lblMsg15").Caption = PgDadosFinanceiroFatura(id).ObsBol2
'                .Item("lblMsg16").Caption = PgDadosFinanceiroFatura(id).ObsBol3
'                '************************************************************************************************************
'            'End With
'=======================================================================
    ' Definição dos dados do cedente
     Dim CedenteInfo As New CedenteInfo
    'CedenteInfo Cedente = new CedenteInfo()
            
        CedenteInfo.cedente = PgDadosEmpresa(ID_Empresa).Nome
        CedenteInfo.Endereco = PgDadosEmpresa(ID_Empresa).Lgr
        CedenteInfo.CNPJ = PgDadosEmpresa(ID_Empresa).CNPJ
        CedenteInfo.banco = pgDadosBanco(PgDadosFinanceiroFatura(idFatura).IdBanco).Numero    '"237"
        CedenteInfo.agencia = pgDadosConta(PgDadosFinanceiroFatura(idFatura).idConta).agencia & "-" & _
        pgDadosConta(PgDadosFinanceiroFatura(idFatura).idConta).AgenciaDV  '"999-7"
        
        CedenteInfo.conta = pgDadosConta(PgDadosFinanceiroFatura(idFatura).idConta).conta & "-" & _
        pgDadosConta(PgDadosFinanceiroFatura(idFatura).idConta).ContaDV  '"999999-7"
        
        CedenteInfo.carteira = "17" 'pgDadosConta(PgDadosFinanceiroFatura(idFatura).idConta).carteira  '"18"
        CedenteInfo.Modalidade = "19" ' pgDadosConta(PgDadosFinanceiroFatura(idFatura).idConta).Variacao    '"19"
        CedenteInfo.Convenio = pgDadosConta(PgDadosFinanceiroFatura(idFatura).idConta).Convenio  '"123456"    ' ATENÇÃO: Alguns Bancos usam um código de convenio para remapear a conta do clientes
        'CedenteInfo.CodCedente = pgDadosConta(PgDadosFinanceiroFatura(idFatura).idConta).Convenio  '"123456"  ' outros bancos chama isto de Codigo do Cedente ou Código do Cliente
        ' outros usam os 2 campos para controles distintos!
        ' Veja com atenção qual é o seu caso e qual destas variáveis deve ser usadas!
        ' Olhe sempre os exemplos em ASP.Net se tiver dúvidas, pois lá há um exemplo para cada banco
        CedenteInfo.UsoBanco = "123"
        ' CedenteInfo.CIP = "456" ' se for informado esse campo o layout muda um pouco

        ' Definição dos dados do sacado
        Dim Sacado As New SacadoInfo
        'SacadoInfo Sacado = new SacadoInfo()
        Sacado.Sacado = PgDadosFinanceiroFatura(idFatura).Sacado
        Sacado.Documento = PgDadosFinanceiroFatura(idFatura).CNPJSacado
        Sacado.Endereco = PgDadosFinanceiroFatura(idFatura).Sacado
        Sacado.Cidade = " " '"São Paulo"
        Sacado.Bairro = " " ' "Centro"
        Sacado.CEP = " " ' "12345-123"
        Sacado.UF = " " '"SP"
        Sacado.Avalista = " " '"Banco XPTO - CNPJ: 123.456.789/00001-23"

        ' Definição das Variáveis do boleto
        'BoletoInfo Boleto = new BoletoInfo()
        Dim Boleto As New BoletoInfo
        Boleto.NossoNumero = PgDadosFinanceiroFatura(idFatura).NossoNumero
        Boleto.NumeroDocumento = PgDadosFinanceiroFatura(idFatura).NossoNumero
        
        Boleto.ParcelaNumero = 1 '2
        Boleto.ParcelaTotal = 1 ' 6
        Boleto.Quantidade = 1 '5
        Boleto.ValorUnitario = PgDadosFinanceiroFatura(idFatura).vlDuplicata ' 20
        Boleto.ValorDocumento = Boleto.Quantidade * Boleto.ValorUnitario
        Boleto.DataDocumento = DateTime.Now
        Dim dVenc As Date
        
        Boleto.DataVencimento = Date - 30 'DateTime.Now.AddDays(-30)
        'Boleto.Especie = "DM" 'Especies.RC
        Boleto.DataDocumento = Now 'DateTime.Now.AddDays(-2)     ' Por padrão é  a data atual, geralmente é a data em que foi feita a compra/pedido, antes de ser gerado o boleto para pagamento
        Boleto.DataProcessamento = Now 'DateTime.Now.AddDays(-1) ' Por padrão é a data atual, pode ser usado como a data em que foi impresso o boleto
        
        ' http:'calculoexato.com.br/parprima.aspx?codMenu=DividBoletoVencido
        ' Se for especificado o valor da mora, este será usado da forma mais simples
        'Boleto.ValorMora = 0.03
        Boleto.PercentualMulta = 0.02 ' 2.0% no mês
        Boleto.PercentualMora = 0.001 ' 0.1% valor percentual ao dia...
        ' Valor original: R$100,00
        ' Valor da multa de 2%: R$2,00
        ' Valor com multa: R$102,00
        ' Valor da 0.1% ao dia por 60 dias (6,00%): R$6,12
        ' Valor com mora: R$108,12
        ' Valor a ser pago: R$108,12
        ' veja também a mesma conta em: http:'exame.abril.com.br/seu-dinheiro/ferramentas/boleto-vencido.shtml
        ' No valor do percentual Mora pode ser usado um valor mensal do tipo:
        ' Boleto.PercentualMora = 0.03 / 30d ' 3% ao mês

        ' Se for especificado a data de pagamento esta será usada como base para o calculo do numero de dias em que será pago
        Boleto.DataPagamento = Now 'Boleto.DataVencimento.AddDays(60)
        
        ' Ativa o calculo de Juros+Mora
        Boleto.CalculaMultaMora = True

        ' Outros valores opcionais
        'Boleto.ValorDesconto = 10
        'Boleto.DataDesconto = DateTime.Now.AddDays(-10)
        'Boleto.ValorAcrescimo = 3
        'Boleto.ValorOutras = 12.34
        Boleto.Instrucoes = "Todas as informações deste bloqueto são de exclusiva responsabilidade do cedente"

        'BoletoInfo Boleto = new BoletoInfo()
        ' O tipo de documento pode ser selecionado para cada boleto, o padrão é DM
        'Boleto.Especie = "DM" 'Especies.DS

        ' Personaliza o boleto com seu logo
        'bltFrm.Boleto.CedenteLogo = BoletoForm2.Properties.Resources.SeuLogo

    
    
        Dim blt As New Boleto
        MsgBox blt.carteira
        ' monta o boleto com os dados específicos nas classes
      blt.MakeBoleto CedenteInfo, Sacado, Boleto
      blt.Save App.Path & "\boleto.bmp"
        MsgBox "boleto gerado em " & App.Path
        ' É possivel também customizar a linha referente o local de pagamento:
      ' Boleto.LocalPagamento = "Pague Preferencialmente no BANCO NOSSA CAIXA S.A. ou na rede bancária até o vencimento"

        ' Configura campos especiais extras no boleto
       ' PrintRecibo (bltFrm)
'           Dim ret As New LayoutBancos
'
'           ret.Init CedenteInfo
'
'           ret.Add Boleto, Sacado
'
'           MsgBox ret.Remessa
           Exit Sub
'trterro:
'           MsgBox Err.Number & " - " & Err.Description
'
End Sub

Private Sub Command3_Click()
    Dim l As String
    l = App.Path & "\boleto.bmp"
    'Picture1.Picture = LoadPicture(l)
End Sub

