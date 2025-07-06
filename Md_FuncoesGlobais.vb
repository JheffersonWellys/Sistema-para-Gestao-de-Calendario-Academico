'---------------------------------------------------------------------------------------
' Módulo   : Md_FuncoesGlobais
' Finalidade: Contém funções auxiliares utilizadas pela Ribbon do GCA, como navegação entre
'             abas e placeholders de funcionalidades ainda não implementadas.
'
' Autor    : H4rzel
' Criado em: 06/07/2025
' Última atualização: 07/07/2025
'---------------------------------------------------------------------------------------

Public Module Md_FuncoesGlobais

#Region "Funções do Menu Ribbon"

    ''' <summary>
    ''' Navega para a aba especificada na Ribbon do GCA.
    ''' Este método atualiza a aba atual e invalida a interface da Ribbon para refletir a mudança.
    ''' É utilizado para alternar entre diferentes seções do GCA, como cronograma, calendário, etc.
    ''' </summary>
    ''' <param name="tabId"></param>
    Public Sub NavegarParaTab(tabId As String)
        currentTab = tabId
        RbbnUI_GCA?.Invalidate()
    End Sub

    ''' <summary>
    ''' Exibe uma mensagem informando que a funcionalidade está em breve disponível.
    ''' Este método é utilizado como um placeholder para funcionalidades que ainda não foram implementadas.
    ''' Ao ser chamado, exibe uma caixa de mensagem com o texto "Em breve..." e um ícone de informação.
    ''' </summary>
    Public Sub EmBreve()
        MessageBox.Show("Em breve...", "Aviso!", MessageBoxButtons.OK, MessageBoxIcon.Information)
        RbbnUI_GCA?.Invalidate()
    End Sub

    Public Function VerificarHabilitacaoDoBotao(control As Office.IRibbonControl) As Boolean
        If RegrasDeHabilitacaoDeBotoesDaRibbon.ContainsKey(control.Id) Then
            Return RegrasDeHabilitacaoDeBotoesDaRibbon(control.Id).Invoke()
        Else
            Return True
        End If
    End Function


#End Region

End Module
