'---------------------------------------------------------------------------------------
' Classe   : Rbbn_GCA
' Finalidade: Implementa a Ribbon (Faixa de Opções) personalizada para o GCA
'             (Gerenciador de Controle Acadêmico) no Excel.
'
' Autor    : H4rzel
' Criado em: 06/07/2025
' Última atualização: 07/07/2025
'---------------------------------------------------------------------------------------

Imports System.Drawing

<Runtime.InteropServices.ComVisible(True)>
Public Class Rbbn_GCA
    Implements Office.IRibbonExtensibility

#Region "Variáveis Globais do Menu Ribbon"

    Private ribbonUI As Office.IRibbonUI

#End Region

#Region "Funcionalidades do Menu Ribbon"

    ''' <summary>
    ''' Classe que implementa a Faixa de Opções (Ribbon) do Excel para o GCA (Gerenciador de Controle Acadêmico).
    ''' Esta classe define as ações e visibilidade dos controles na Ribbon, além de carregar a interface a partir de um recurso XML.
    ''' </summary>
    Public Sub New()
    End Sub

    ''' <summary>
    ''' Retorna o XML que define a Faixa de Opções (Ribbon) para o GCA.
    ''' Este método é chamado pelo Excel para carregar a interface da Ribbon.
    ''' O XML é obtido de um recurso incorporado no assembly do Excel, permitindo personalizar a aparência e funcionalidade da Ribbon.
    ''' </summary>
    ''' <param name="ribbonID"></param>
    ''' <returns></returns>
    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("GCA___Turma_000._0000._0000.Rbbn_GCA.xml")
    End Function

#End Region

#Region "Funções da Faixa de Opções da Ribbon"

    ''' <summary>
    ''' Carrega a interface da Faixa de Opções (Ribbon) quando o Excel é iniciado ou quando a Ribbon é atualizada.
    ''' Esta função é chamada automaticamente pelo Excel.
    ''' </summary>
    ''' <param name="ribbonUI"></param>
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbonUI = ribbonUI
        RbbnUI_GCA = Me.ribbonUI
    End Sub

    ''' <summary>
    ''' Determina se uma aba da Faixa de Opções deve estar visível, com base na aba atual selecionada.
    ''' </summary>
    ''' <param name="control">Controle da Ribbon que solicita a visibilidade.</param>
    Public Function VerificarVisibilidadeDaAba(control As Office.IRibbonControl) As Boolean
        Return control.Id = currentTab
    End Function

    ''' <summary>
    ''' Executa a ação associada ao botão clicado na Faixa de Opções (Ribbon).
    ''' </summary>
    ''' <param name="control">Controle da Ribbon que acionou a ação.</param>
    Public Sub ExecutarAcaoDaRibbon(control As Office.IRibbonControl)
        If MapeamentoDeAcoesDaRibbon.ContainsKey(control.Id) Then
            Dim actionEnum = MapeamentoDeAcoesDaRibbon(control.Id)
            If ExecutorDeAcoesDaRibbon.ContainsKey(actionEnum) Then
                ExecutorDeAcoesDaRibbon(actionEnum).Invoke()
            End If
        Else
            NavegarParaTab("Tb_Logon")
        End If
    End Sub

    ''' <summary>
    ''' Retorna o ícone correspondente ao botão da Ribbon com base na ação associada.
    ''' </summary>
    ''' <param name="control">Controle da Ribbon que está solicitando o ícone.</param>
    ''' <returns>Bitmap com o ícone correspondente ou ícone padrão, se não encontrado.</returns>
    Public Function ObterIconeDoBotao(control As Office.IRibbonControl) As Bitmap
        If MapeamentoDeAcoesDaRibbon.ContainsKey(control.Id) Then
            Dim action = MapeamentoDeAcoesDaRibbon(control.Id)
            If MapaDeIconesDeBotoesDaRibbon.ContainsKey(action) Then
                Return MapaDeIconesDeBotoesDaRibbon(action)
            End If
        End If
        Return My.Resources.icn_IconePadrao
    End Function

    ''' <summary>
    ''' Verifica se o botão da Ribbon deve estar habilitado ou desabilitado.
    ''' Esta função é chamada pelo Excel para determinar a habilitação do botão com base no estado do sistema.
    ''' Se as configurações do sistema estiverem completas, o botão será habilitado; caso contrário, será desabilitado.
    ''' </summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function VerificarHabilitacaoDoBotao(control As Office.IRibbonControl) As Boolean
        If RegrasDeHabilitacaoDeBotoesDaRibbon.ContainsKey(control.Id) Then
            Return RegrasDeHabilitacaoDeBotoesDaRibbon(control.Id).Invoke()
        Else
            Return True
        End If
    End Function

#End Region

#Region "Auxiliares"

    ''' <summary>
    ''' Obtém o texto de um recurso incorporado no assembly do Excel.
    ''' Este método é usado para carregar a definição da Faixa de Opções (Ribbon) a partir de um arquivo XML incorporado.
    ''' </summary>
    ''' <param name="resourceName"></param>
    ''' <returns></returns>
    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
