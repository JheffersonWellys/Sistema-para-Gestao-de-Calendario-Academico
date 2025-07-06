'---------------------------------------------------------------------------------------
' Módulo   : Md_ExcelHelper
' Finalidade: Auxilia no acesso à pasta de trabalho, planilhas, intervalos e tabelas do Excel.
'
' Autor    : H4rzel
' Criado em: 06/07/2025
' Última atualização: 07/07/2025
'---------------------------------------------------------------------------------------

Public Module Md_ExcelHelper

#Region "Propriedades para Recuperação Pasta de Trabalho"

    ''' <summary>
    ''' Retorna o objeto Workbook atual do Excel associado ao GCA (Gerenciador de Controle Acadêmico).
    ''' Permite manipular o pasta de trabalho principal da aplicação.
    ''' </summary>
    Public ReadOnly Property GCA As Excel.Workbook
        Get
            Return Globals.ThisWorkbook.InnerObject
        End Get
    End Property


#End Region

#Region "Propriedades para Recuperação Planilhas"

    ''' <summary>
    ''' Retorna a planilha do menu principal do GCA.
    ''' Usada para acessar a planilha que contém o menu inicial.
    ''' </summary>
    Public ReadOnly Property PL_MENU As Excel.Worksheet
        Get
            Return Globals.GCA_PL_MENU?.InnerObject
        End Get
    End Property

    ''' <summary>
    ''' Retorna a planilha de cronograma acadêmico do GCA.
    ''' Utilizada para acessar e manipular o cronograma de aulas.
    ''' </summary>
    Public ReadOnly Property PL_CRONOGRAMA As Excel.Worksheet
        Get
            Return Globals.GCA_PL_CRONOGRAMA?.InnerObject
        End Get
    End Property

    ''' <summary>
    ''' Retorna a planilha de calendário acadêmico do GCA.
    ''' Utilizada para visualizar o calendário acadêmico.
    ''' </summary>
    Public ReadOnly Property PL_CALENDARIO_ACADEMICO As Excel.Worksheet
        Get
            Return Globals.GCA_PL_CALENDARIO_ACADEMICO?.InnerObject
        End Get
    End Property

    ''' <summary>
    ''' Retorna a planilha de configurações do GCA.
    ''' Utilizada para acessar as configurações e preferências do sistema.
    ''' </summary>
    Public ReadOnly Property PL_CONFIGURACOES As Excel.Worksheet
        Get
            Return Globals.GCA_PL_CONFIGURACOES?.InnerObject
        End Get
    End Property

#End Region

#Region "Funções para Recuperação Segura de Objetos"

    ''' <summary>
    ''' Tenta obter um objeto Range de uma planilha pelo nome do intervalo.
    ''' Retorna o objeto Range ou Nothing caso não encontrado ou erro.
    ''' </summary>
    ''' <param name="planilha">Planilha onde o intervalo está localizado.</param>
    ''' <param name="nomeIntervalo">Nome do intervalo (Range) a ser recuperado.</param>
    ''' <returns>Objeto Excel.Range ou Nothing.</returns>
    Public Function ObterIntervalo(planilha As Excel.Worksheet, nomeIntervalo As String) As Excel.Range
        Try
            Dim r As Excel.Range = planilha.Range(nomeIntervalo)
            Return If(r IsNot Nothing, r, Nothing)
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Tenta obter um objeto ListObject (tabela) de uma planilha pelo nome da tabela.
    ''' Retorna o objeto ListObject ou Nothing caso não encontrado ou erro.
    ''' </summary>
    ''' <param name="planilha">Planilha onde a tabela está localizada.</param>
    ''' <param name="nomeTabela">Nome da tabela a ser recuperada.</param>
    ''' <returns>Objeto Excel.ListObject ou Nothing.</returns>
    Public Function ObterTabela(planilha As Excel.Worksheet, nomeTabela As String) As Excel.ListObject
        Try
            Dim tbl As Excel.ListObject = planilha.ListObjects(nomeTabela)
            Return If(tbl IsNot Nothing, tbl, Nothing)
        Catch
            Return Nothing
        End Try
    End Function

#End Region

#Region "Funções para Manipulação de Dados"

#End Region

End Module
