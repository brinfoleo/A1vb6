VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIFormA1 
   AutoShowChildren=   0   'False
   BackColor       =   &H00C0C000&
   Caption         =   "A1 - Aplicativo de Gestão Empresarial"
   ClientHeight    =   5985
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17940
   Icon            =   "MDIFormA1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   180
      Top             =   4260
   End
   Begin ComCtl3.CoolBar cbMenu 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17940
      _ExtentX        =   31644
      _ExtentY        =   1535
      _CBWidth        =   17940
      _CBHeight       =   870
      _Version        =   "6.0.8169"
      Child1          =   "tbMenuComercial"
      MinHeight1      =   810
      Width1          =   2130
      NewRow1         =   0   'False
      Child2          =   "tbMenuFaturamento"
      MinHeight2      =   810
      Width2          =   3720
      NewRow2         =   0   'False
      Child3          =   "tbMenuFinanceiro"
      MinHeight3      =   810
      Width3          =   1635
      NewRow3         =   0   'False
      Begin MSComctlLib.Toolbar tbMenuFinanceiro 
         Height          =   810
         Left            =   6075
         TabIndex        =   4
         Top             =   30
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   1429
         ButtonWidth     =   1455
         ButtonHeight    =   1429
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Gerenciador Financeiro"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Contas Pagar / Receber"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Baixa Automatica"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbMenuFaturamento 
         Height          =   810
         Left            =   2325
         TabIndex        =   3
         Top             =   30
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   1429
         ButtonWidth     =   1455
         ButtonHeight    =   1429
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Gerenciador de NFe"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Emissão de NFe"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Carta correção"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cancelar NFe"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbMenuComercial 
         Height          =   810
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1429
         ButtonWidth     =   1455
         ButtonHeight    =   1429
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pré-Venda"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Gerenciador de Vendas"
               ImageIndex      =   1
            EndProperty
         EndProperty
         Begin VB.Timer tmConexaoServidor 
            Interval        =   30000
            Left            =   9660
            Top             =   360
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   240
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFormA1.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFormA1.frx":1DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFormA1.frx":2C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFormA1.frx":4958
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFormA1.frx":6632
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFormA1.frx":6D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFormA1.frx":7606
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFormA1.frx":92E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFormA1.frx":AFBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFormA1.frx":CC94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar BarraStatus 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   5490
      Width           =   17940
      _ExtentX        =   31644
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18830
            MinWidth        =   71
            Text            =   "1"
            TextSave        =   "1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "2"
            TextSave        =   "2"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   3995
            MinWidth        =   71
            Text            =   "Status do Servidor: Conectado"
            TextSave        =   "Status do Servidor: Conectado"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   71
            Text            =   "0000.0.0"
            TextSave        =   "0000.0.0"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1693
            MinWidth        =   71
            TextSave        =   "03/05/2021"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   873
            MinWidth        =   71
            TextSave        =   "23:28"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsMain 
      Left            =   180
      Top             =   3780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu Empresa 
      Caption         =   "Empresa"
      Begin VB.Menu EmpresaCadastro 
         Caption         =   "Cadastro"
      End
      Begin VB.Menu ExportarDados 
         Caption         =   "Exportar Dados"
      End
   End
   Begin VB.Menu Clientes 
      Caption         =   "Clientes"
      Begin VB.Menu ClientesCadastro 
         Caption         =   "Cadastro"
      End
      Begin VB.Menu ClientesListagem 
         Caption         =   "Listagem"
      End
   End
   Begin VB.Menu Fornecedores 
      Caption         =   "Fornecedores"
      Begin VB.Menu FornecedoresCadastro 
         Caption         =   "Cadastro"
      End
   End
   Begin VB.Menu Transportadoras 
      Caption         =   "Transportadoras"
      Begin VB.Menu TransportadorasCadastro 
         Caption         =   "Cadastro"
      End
   End
   Begin VB.Menu financeiro 
      Caption         =   "Financeiro"
      Begin VB.Menu FinanceiroCadastro 
         Caption         =   "Cadastro"
         Begin VB.Menu financeiroContasPagarReceber 
            Caption         =   "Contas a Pagar/Receber"
         End
         Begin VB.Menu FinanceiroContasPagarReceberFixas 
            Caption         =   "Contas a Pagar/Receber Fixas"
         End
         Begin VB.Menu FinancSp 
            Caption         =   "-"
         End
         Begin VB.Menu FinanceiroCondicoesPagamento 
            Caption         =   "Condições de Pagamento"
         End
         Begin VB.Menu FinanceiroTipoDocumento 
            Caption         =   "Tipo de Documento"
         End
         Begin VB.Menu FinanceiroBanco 
            Caption         =   "Banco"
         End
         Begin VB.Menu FinanceiroContaCorrente 
            Caption         =   "Conta"
         End
         Begin VB.Menu PlanoContas 
            Caption         =   "Plano de Contas"
         End
         Begin VB.Menu FinanceiroCentroCustos 
            Caption         =   "Centro de Custos"
         End
      End
      Begin VB.Menu FinanceiroContasPR 
         Caption         =   "Gerenciador Financeiro"
      End
      Begin VB.Menu FinanceiroBxAutomatica 
         Caption         =   "Baixa automatica de titulos"
      End
      Begin VB.Menu FinanceiroContaMov 
         Caption         =   "Livro Caixa"
         Begin VB.Menu FinanceiroLancamentoConta 
            Caption         =   "Lançamento de Conta"
         End
         Begin VB.Menu FinanceiroExtratoConta 
            Caption         =   "Extrato de Conta"
         End
      End
      Begin VB.Menu financeiroContasPagarReceberDRE 
         Caption         =   "DRE"
      End
      Begin VB.Menu financeiroCnab240 
         Caption         =   "Gerar CNAB 240"
      End
   End
   Begin VB.Menu Estoque 
      Caption         =   "Estoque"
      Begin VB.Menu EstoqueCadastro 
         Caption         =   "Cadastro"
         Begin VB.Menu EstoqueDeposito 
            Caption         =   "Depósito"
         End
         Begin VB.Menu EstoqueProduto 
            Caption         =   "Produto"
         End
         Begin VB.Menu EstoqueSPC001 
            Caption         =   "-"
         End
         Begin VB.Menu EstoqueGrupoProdutos 
            Caption         =   "Grupo de Produtos"
         End
         Begin VB.Menu EstoqueSubGrupo 
            Caption         =   "Subgrupo"
         End
         Begin VB.Menu Fabricante 
            Caption         =   "Fabricante"
         End
         Begin VB.Menu EstoqueMovimento 
            Caption         =   "Movimento"
         End
         Begin VB.Menu EstoqueUnidadeMedidas 
            Caption         =   "Unidades de Medidas"
         End
      End
      Begin VB.Menu EstoqueGerenciador 
         Caption         =   "Gerenciador de Estoque"
      End
      Begin VB.Menu EstoqueManutencao 
         Caption         =   "Manutencao"
      End
      Begin VB.Menu EstoqueKardex 
         Caption         =   "Kardex"
      End
      Begin VB.Menu EstoqueRelatorio 
         Caption         =   "Relatorio"
      End
      Begin VB.Menu EstoqueSPC002 
         Caption         =   "-"
      End
      Begin VB.Menu EstoquePedidoCompra 
         Caption         =   "Pedido de Compra"
      End
   End
   Begin VB.Menu Faturamento 
      Caption         =   "Faturamento"
      Begin VB.Menu mnContrato 
         Caption         =   "Contrato"
         Begin VB.Menu mnuContratoGerenciador 
            Caption         =   "Gerenciador de Contratos"
         End
      End
      Begin VB.Menu FaturamentoPV 
         Caption         =   "Pré Venda"
         Begin VB.Menu emissao 
            Caption         =   "Emissão"
         End
         Begin VB.Menu Gerenciador 
            Caption         =   "Gerenciador de Pré-Venda"
         End
         Begin VB.Menu GerarPreVendaemLote 
            Caption         =   "Gerar Pré-venda em Lote"
         End
         Begin VB.Menu FaturamentoPVRelatorio 
            Caption         =   "Relatório"
         End
      End
      Begin VB.Menu FaturamentoNotaFiscal 
         Caption         =   "Nota Fiscal"
         Begin VB.Menu GerenciadorNFe 
            Caption         =   "Gerenciador de NF-e"
         End
         Begin VB.Menu FaturamentoCC 
            Caption         =   "Carta de Correção"
         End
         Begin VB.Menu CancelarNFe 
            Caption         =   "Cancelar Nf-e"
         End
         Begin VB.Menu InutilizarNumeroNF 
            Caption         =   "Inutilizar Numero NF"
         End
         Begin VB.Menu Sep 
            Caption         =   "-"
         End
         Begin VB.Menu NFEntrada 
            Caption         =   "Entrada"
            Begin VB.Menu FaturamentoNFeEntrada 
               Caption         =   "Cadastro de Nota Fiscal"
            End
         End
         Begin VB.Menu NFSaida 
            Caption         =   "Saida"
            Begin VB.Menu FaturamentoNotaFiscalEmissaoNFe 
               Caption         =   "Emissao de NF-e"
            End
            Begin VB.Menu NFConsulta 
               Caption         =   "Consultar NFe"
            End
         End
         Begin VB.Menu FaturamentoRelatorio 
            Caption         =   "Relatorio"
            Begin VB.Menu RptFatAnaliseVendas 
               Caption         =   "Analise de Faturas/NFe"
            End
         End
         Begin VB.Menu FaturamentoNotaFiscalTipoNotaFiscal 
            Caption         =   "Tipo de Nota Fiscal"
         End
      End
   End
   Begin VB.Menu RecursosHumanos 
      Caption         =   "Departamento Pessoal"
      Begin VB.Menu RHFuncionarios 
         Caption         =   "Funcionarios"
      End
      Begin VB.Menu RHCargo 
         Caption         =   "Cargos e Funções"
      End
      Begin VB.Menu FolhaPagamento 
         Caption         =   "Folha de Pagamento"
         Begin VB.Menu RHExtratoPagamento 
            Caption         =   "Folha de Pagamento"
         End
         Begin VB.Menu RHComissao 
            Caption         =   "Ajustar/Calcular Comissão"
         End
      End
      Begin VB.Menu RHFolhaPonto 
         Caption         =   "Folha de Ponto"
      End
   End
   Begin VB.Menu Tributos 
      Caption         =   "Tributos"
      Begin VB.Menu GerencTributos 
         Caption         =   "Gerenciador de Tributos"
      End
      Begin VB.Menu TibutosCadastro 
         Caption         =   "Cadastro"
         Begin VB.Menu TributosPais 
            Caption         =   "Codigo Pais"
         End
         Begin VB.Menu TributosUFICMS 
            Caption         =   "Codigo UF e ICMS"
         End
         Begin VB.Menu TributosMunicipio 
            Caption         =   "Codigo do Municipio"
         End
         Begin VB.Menu TributosCST 
            Caption         =   "CST - Codigo de Situação Fiscal"
         End
         Begin VB.Menu TributosNCM 
            Caption         =   "NCM - Nomeclatura Comum no Mercosul"
         End
         Begin VB.Menu TributosCFOP 
            Caption         =   "CFOPs"
         End
      End
   End
   Begin VB.Menu Sistema 
      Caption         =   "Sistema"
      WindowList      =   -1  'True
      Begin VB.Menu Configuracoes 
         Caption         =   "Configurações"
      End
      Begin VB.Menu SistemaVizualizarBasedeDados 
         Caption         =   "Vizualizar Base de Dados"
         Enabled         =   0   'False
      End
      Begin VB.Menu SisConexao 
         Caption         =   "Conexão"
         Begin VB.Menu GerencConexao 
            Caption         =   "Usuarios Conectados"
         End
         Begin VB.Menu ConexaoBD 
            Caption         =   "Conexão com a Base de Dados"
         End
      End
      Begin VB.Menu sisBancoDados 
         Caption         =   "Banco de Dados"
         Begin VB.Menu AtualizacaodaBasedeDados 
            Caption         =   "Atualização da Base de Dados"
         End
         Begin VB.Menu ConexaoBasedeDados 
            Caption         =   "Conexão a Base de Dados"
         End
         Begin VB.Menu SistemaImpExp 
            Caption         =   "Importar e Exportar Dados"
         End
         Begin VB.Menu GerarBackup 
            Caption         =   "Gerar/Recuperar Backup"
         End
      End
      Begin VB.Menu Usuario 
         Caption         =   "Usuario / Grupo de Usuarios"
      End
      Begin VB.Menu Chat 
         Caption         =   "Chat"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Ajuda 
      Caption         =   "Ajuda"
      Begin VB.Menu Sobre 
         Caption         =   "Sobre o A1"
      End
      Begin VB.Menu SuporteOnLine 
         Caption         =   "Suporte On Line"
      End
      Begin VB.Menu AjustesdoLeo 
         Caption         =   "Ajustes do Sistema"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "MDIFormA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub AjustesdoLeo_Click()
    Dim Xpws As String
    
    Xpws = InputBox("Senha", "Senha")
    If Xpws <> "7821" Then Exit Sub
    
    '***********************************************************
    '**** AJUSTAR CLASSIFICAÇÃO DAS NFe
    If MsgBox("Ajustar MovFisco / MovFinanceir", vbYesNo, App.EXEName) = vbYes Then
        BD.Execute "UPDATE faturamentonfe SET movfisco = 1"
        BD.Execute "UPDATE faturamentonfe SET movfinanceiro = 1"
        BD.Execute "UPDATE faturamentonfe SET enviorf = 1"
        BD.Execute "UPDATE faturamentonfe SET impfatura = 1"
        MsgBox "FaturamentoNFe - atualizado", vbInformation, App.EXEName
    End If
    
    '***********************************************************
    '**** AJUSTAR CLASSIFICAÇÃO DO ESTOQUE
    If MsgBox("ESTOQUE: Montar referencia do produto com base na descricao do material?", vbYesNo, App.EXEName) = vbYes Then
        Dim sSQL As String
        Dim Rst As Recordset
        
        Dim c As Integer
        Dim a As Integer
        Dim total As Integer
        Dim sCod As String
        
        sSQL = "SELECT * FROM estoqueproduto WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito
        Set Rst = RegistroBuscar(sSQL)
        If Rst.BOF And Rst.EOF Then
            Else
                Rst.MoveFirst
                total = Rst.RecordCount
                c = 0
                a = 0
                Do Until Rst.EOF
                DoEvents
                    If InStr(Rst.Fields("Descricao"), """") <> 0 Then
                    'txtInformacoesComplementares.Enabled = True
                    'txtInformacoesComplementares.Text = txtInformacoesComplementares.Text & vbCrLf & Rst.fields("Descricao")
                    sCod = MontarReferenciaProduto(Rst.Fields("id"))
                    sSQL = "UPDATE estoqueproduto SET Referencia = " & sCod & " WHERE id=" & Rst.Fields("id")
                    BD.Execute sSQL
                    c = c + 1
                    
                End If
                a = a + 1
                Me.Caption = "Atual: " & a & " Modificados: " & c & " Total: " & total
                Rst.MoveNext
                Loop
        End If
        Rst.Close
    End If

End Sub


Private Sub Chat_Click()
    formChat.Show
End Sub

Private Sub ClientesListagem_Click()
    formClientesRelatorio.Show
    
End Sub

Private Sub ConexaoBasedeDados_Click()
    formConexao.Show 1
End Sub

Private Sub EstoqueGerenciador_Click()
    formEstoqueGerenciador.Show
End Sub
Private Sub EstoqueRelatorio_Click()
    formEstoqueAnalise.Show
End Sub
Private Sub ExportarDados_Click()
    formEmpresaExportarDados.Show
End Sub
Private Sub Fabricante_Click()
    formEstoqueFabricante.Show
End Sub
Private Sub FaturamentoCC_Click()
    formFaturamentoNFeCartaCorrecao.Show
End Sub

Private Sub FaturamentoPVRelatorio_Click()
    formFaturamentoPVRelatorios.Show
End Sub

Private Sub FinanceiroBxAutomatica_Click()
    formFinanceiroBaixaAutomaticaTitulo.Show
End Sub

Private Sub financeiroCnab240_Click()
    formFinanceiroCnab240.Show
End Sub

Private Sub financeiroContasPagarReceberDRE_Click()
    formFinanceiroDRE.Show
    
End Sub

Private Sub FinanceiroContasPagarReceberFixas_Click()
    formFinanceiroContasPRFixa.Show
End Sub
Private Sub FinanceiroExtratoConta_Click()
    formFinanceiroContaExtrato.Show
End Sub

Private Sub FinanceiroLancamentoConta_Click()
    formFinanceiroContaMov.Show
    
End Sub

Private Sub GerarPreVendaemLote_Click()
    formFaturamentoPvLote.Show
End Sub

Private Sub GerencConexao_Click()
    formUsuConexaoGerenciador.Show
End Sub
Private Sub InutilizarNumeroNF_Click()
    formFaturamentoNFeInutilizarNumero.Show
End Sub

Private Sub MDIForm_Activate()
  
    With BarraStatus
        .Panels(1).Text = PgDadosEmpresa(ID_Empresa).Nome
        .Panels(4).Text = App.Major & "." & App.Minor & "." & App.Revision
        '.Panels(3).Text = "XXXXX"
        .Panels(2).Text = UCase(PgDadosUsuario(ID_Usuario).Nome)
        .Panels(2).Alignment = sbrLeft
    End With
End Sub

Private Sub mnuContratoGerenciador_Click()
    formContratoGerenciador.Show
End Sub


Private Sub NFConsulta_Click()
    formFaturamentoNFeConsulta.Show
End Sub
Private Sub PlanoContas_Click()
    formFinanceiroPlanoContas.Show
End Sub

Private Sub RHExtratoPagamento_Click()
    formRHFuncionarioFolhadePagamento.Show
End Sub

Private Sub RHFolhaPonto_Click()
    formRHFuncionarioFolhaPonto.Show
End Sub
Private Sub RptFatAnaliseVendas_Click()
    formFaturamentoAnalise.Show
End Sub
Private Sub SuporteOnLine_Click()


 
'    On Error Resume Next
'    Dim caminho As String
'    caminho = PgDadosConfig.pFileArmazenamento & "\Suporte\Suporte.exe"
'    If Trim(Dir(caminho, vbArchive)) <> "" Then
'        If MsgBox("Deseja acionar o suporte?", vbYesNo + vbInformation, App.EXEName) = vbYes Then
'            Shell caminho, vbNormalFocus
'        End If
'    Else
'        MsgBox "Sistema de suporte ausente!", vbCritical, App.EXEName
'    End If
End Sub
Private Sub tbMenuComercial_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case tbMenuComercial.Buttons(Button.Index).ToolTipText
        Case "Pré-Venda"
            formFaturamentoPV.Show
        Case "Gerenciador de Vendas"
            formComercialGerenciadorClientes.Show
    End Select
End Sub
Private Sub AtualizacaodaBasedeDados_Click()
    formBaseDadosAtualizar.Show
End Sub
Private Sub CancelarNFe_Click()
    formFaturamentoNFeCancelada.Show
End Sub
Private Sub ClientesCadastro_Click()
    formClientes.Show
End Sub
Private Sub ConexaoBD_Click()
    formConexao.Show 1
End Sub

Private Sub Configuracoes_Click()
    formConfiguracoes.Show
End Sub
Private Sub emissao_Click()
    formFaturamentoPV.Show
End Sub
Private Sub EmpresaCadastro_Click()
    formEmpresas.Show
End Sub
Private Sub EstoqueDeposito_Click()
    formEstoqueDeposito.Show
    
End Sub
Private Sub EstoqueGrupoProdutos_Click()
    formEstoqueGrupos.Show
End Sub
Private Sub EstoqueKardex_Click()
    formEstoqueKardex.Show
End Sub
Private Sub EstoqueManutencao_Click()
    formEstoqueManutencao.Show
End Sub
Private Sub EstoqueMovimento_Click()
    formEstoqueMovimento.Show
End Sub
Private Sub EstoquePedidoCompra_Click()
    formEstoquePedidoCompra.Show
End Sub

Private Sub EstoqueProduto_Click()
    formEstoqueProduto.Show
End Sub

Private Sub EstoqueSubGrupo_Click()
    formEstoqueSubGrupo.Show
End Sub

Private Sub EstoqueUnidadeMedidas_Click()
    formEstoqueUnidadeMedida.Show
End Sub


Private Sub FaturamentoNFeEntrada_Click()
    formFaturamentoNFeEntrada.Show
End Sub

Private Sub FaturamentoNotaFiscalEmissaoNFe_Click()
    formFaturamentoNFe.Show
End Sub

Private Sub FaturamentoNotaFiscalTipoNotaFiscal_Click()
    formFaturamentoTipoNotaFiscal.Show
End Sub

Private Sub FinanceiroBanco_Click()
    formFinanceiroBancoCadastro.Show
End Sub

Private Sub FinanceiroCentroCustos_Click()
    formFinanceiroCentroCustos.Show
End Sub

Private Sub FinanceiroCondicoesPagamento_Click()
    formFinanceiroCondicoesPagamento.Show
End Sub

Private Sub FinanceiroContaCorrente_Click()
    formFinanceiroConta.Show
End Sub

Private Sub FinanceiroContasPagarReceber_Click()
    formFinanceiroContasPRCadastro.Show
End Sub

Private Sub FinanceiroContasPR_Click()
    formFinanceiroContasPRGerenciador.Show
End Sub

Private Sub FinanceiroTipoDocumento_Click()
    formFinanceiroTipoDocumento.Show
    
End Sub



Private Sub FornecedoresCadastro_Click()
    formFornecedores.Show
End Sub


Private Sub GerarBackup_Click()
    formBackup.Show 1
End Sub

Private Sub Gerenciador_Click()
    formComercialGerenciadorClientes.Show
End Sub

Private Sub GerenciadorNFe_Click()
    formFaturamentoNFeGerenciador.Show
End Sub

Private Sub GerencTributos_Click()
    formTributacaoGerenciador.Show
End Sub

Private Sub MDIForm_Load()
    Dim Vis     As Boolean
    Dim Menu    As String
    Dim i       As Integer
'
'    With BarraStatus
'        .Panels(1).Text = PgDadosEmpresa(ID_Empresa).Nome
'        .Panels(4).Text = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
'        '.Panels(3).Text = "XXXXX"
'        .Panels(2).Text = UCase(PgDadosUsuario(ID_Usuario).Nome)
'        .Panels(2).Alignment = sbrLeft
'    End With
    
    
    '##########################################################################
    '### MOSTRAS OS MENUS DO USUARIO
    Menu = PgDadosUsuario(ID_Usuario).Menus
    For i = 1 To Len(Trim(Menu))
        Vis = IIf(Mid(Menu, i, 1) = 1, True, False)
        cbMenu.Bands(i).Visible = Vis
    Next
    '##########################################################################
    
    
   
End Sub

Private Sub MDIForm_Terminate()
    FinalizandoSistema
End Sub

Private Sub RHCargo_Click()
    formRHFuncionarioCargo.Show
End Sub

Private Sub RHComissao_Click()
    formRHFuncionarioComissao.Show
End Sub

Private Sub RHFuncionarios_Click()
    formRHFuncionarioCadastro.Show
End Sub

Private Sub SistemaImpExp_Click()
    formImpExpDados.Show
End Sub

Private Sub SistemaVizualizarBasedeDados_Click()
    formdbDBGrid.Show
End Sub

Private Sub Sobre_Click()
'    formChat.Show

    formSobre.Show
'    Dim vReg(100) As Variant
'   Dim Rst As Recordset
'   Dim sSQL As String
'    Dim Id As Integer
'    sSQL = "SELECT * FROM FinanceirocontasPRCadastro"
'    Set Rst = RegistroBuscar(sSQL)
'    Rst.MoveFirst
'    Do Until Rst.EOF
'        Id = Rst.Fields("id")
'        vReg(0) = Array("VlDuplicata", ChkVal(Rst.Fields("VlDuplicata"), 0, 2), "S")
'        RegistroAlterar "FinanceiroContasPRCadastro", vReg, 0, "id = " & Id
'        Rst.MoveNext
'    Loop
'    MsgBox "terminou"
End Sub



Private Sub tbMenuFaturamento_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case tbMenuFaturamento.Buttons(Button.Index).ToolTipText
        Case "Gerenciador de NFe"
            formFaturamentoNFeGerenciador.Show
        Case "Emissão de NFe"
            formFaturamentoNFe.Show
        Case "Cancelar NFe"
            formFaturamentoNFeCancelada.Show
        Case "Carta correção"
            formFaturamentoNFeCartaCorrecao.Show
    End Select

End Sub

Private Sub tbMenuFinanceiro_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case tbMenuFinanceiro.Buttons(Button.Index).ToolTipText
        Case "Gerenciador Financeiro"
            formFinanceiroContasPRGerenciador.Show
        Case "Contas Pagar / Receber"
            formFinanceiroContasPRCadastro.Show
        Case "Baixa Automatica"
            formFinanceiroBaixaAutomaticaTitulo.Show
    End Select
End Sub

Private Sub Timer1_Timer()
    'updateSistema
    chkUsuariosConectado
    
End Sub

Private Sub tmConexaoServidor_Timer()
    If TestarConexaoServidor = True Then
            BarraStatus.Panels(3).Text = "Status do Servidor: Conectado"
        Else
            BarraStatus.Panels(3).Text = "Status do Servidor: Desconectado"
    End If

End Sub

Private Sub TransportadorasCadastro_Click()
    formTransportadoras.Show
End Sub





Private Sub TributosCFOP_Click()
    formTributacaoCFOP.Show
End Sub

Private Sub TributosCST_Click()
    formTributacaoCST.Show
End Sub

Private Sub TributosMunicipio_Click()
    formTributacaoMunicipio.Show
End Sub

Private Sub TributosNCM_Click()
    formTributacaoNCM.Show
End Sub

Private Sub TributosPais_Click()
    formTributacaoPais.Show
End Sub

Private Sub TributosUFICMS_Click()
    formTributacaoUF.Show
End Sub



Private Sub Usuario_Click()
    formUsuGerenciador.Show
End Sub
