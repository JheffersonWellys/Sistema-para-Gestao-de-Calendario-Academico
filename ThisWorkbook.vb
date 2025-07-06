
Public Class ThisWorkbook

    ''' <summary>
    ''' Evento disparado quando o Workbook é aberto ou o Excel é iniciado.
    ''' Este método pode ser usado para inicializar variáveis, carregar dados ou ativar a Faixa de Opções personalizada.
    ''' </summary>
    Private Sub ThisWorkbook_Startup() Handles Me.Startup

        PL_MENU.Select()

    End Sub

    ''' <summary>
    ''' Evento disparado quando o Excel é fechado ou o Workbook é fechado.
    ''' Este método pode ser usado para liberar recursos ou salvar dados antes do fechamento.
    ''' </summary>
    Private Sub ThisWorkbook_Shutdown() Handles Me.Shutdown
    End Sub

    ''' <summary>
    ''' Cria o objeto IRibbonExtensibility para a Faixa de Opções personalizada.
    ''' Este método é chamado pelo Excel para inicializar a Faixa de Opções.
    ''' </summary>
    ''' <returns></returns>
    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Rbbn_GCA()
    End Function

End Class
