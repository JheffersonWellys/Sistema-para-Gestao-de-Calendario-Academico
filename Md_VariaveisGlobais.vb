'---------------------------------------------------------------------------------------
' Módulo   : Md_VariaveisGlobais
' Finalidade: Armazena variáveis globais utilizadas pela Ribbon do GCA (Gerenciador de Controle Acadêmico),
'             incluindo identificadores, mapeamento de ações, ícones e controle da aba ativa.
'
' Autor    : H4rzel
' Criado em: 06/07/2025
' Última atualização: 07/07/2025
'---------------------------------------------------------------------------------------

Imports System.Drawing

Public Module Md_VariaveisGlobais

#Region "Office"

    Public RbbnUI_GCA As Office.IRibbonUI

#End Region

#Region "Strings"

    Public currentTab As String = "Tb_Logon"

#End Region

#Region "Enums"

    Public Enum IdentificadoresDeBotoesDaRibbon
        IniciarSessao
        FinalizarSessao
        IrPara_Configuracoes_Administrativas
        IrPara_Configuracoes_Infraestrutura
        IrPara_Configuracoes_Academicas
        IrPara_Configuracoes_Educacionais
        IrPara_Configuracoes_Eventos
        IrPara_Edicao_Cronograma
        Configuracoes_Administrativa_VoltarPara_MenuInicial
        Configurar_TurnoLetivo
        Configurar_HorariosDeAula
        Configurar_TiposUnidadesCurriculares
        Configuracoes_Infraestrutura_VoltarPara_MenuInicial
        Configurar_Blocos
        Configurar_Andares
        Configurar_SalasDeAula
        Configuracoes_Academicas_VoltarPara_MenuInicial
        Configurar_Docentes
        Configurar_AutorizacoesParaLecionar
        Configurar_Atestados
        Configuracoes_Educacionais_VoltarPara_MenuInicial
        Configurar_UnidadeEducacional
        Configurar_AreaProfissional
        Configurar_NomeDoCurso
        Configurar_UnidadesCurriculares
        Configurar_CodigoDaTurma
        Configuracoes_Eventos_VoltarPara_MenuInicial
        Configurar_Feriados
        Configurar_Recessos
        Configurar_DatasEventuais
        Edicao_Cronograma_VoltarPara_MenuInicial
        CriarCronograma_ComIa
        EditarCronograma_Manualmente
        EditarCronograma_ComIa
        Visualizar_Erros
        Visualizar_CalendarioAcademico
        ExportarEm_PDF
        ExportarEm_XLSX
    End Enum

#End Region

#Region "Dictionarys"

    Public ReadOnly MapeamentoDeAcoesDaRibbon As New Dictionary(Of String, IdentificadoresDeBotoesDaRibbon) From {
        {"Bttn_IniciarSessao", IdentificadoresDeBotoesDaRibbon.IniciarSessao},
        {"Bttn_FinalizarSessao", IdentificadoresDeBotoesDaRibbon.FinalizarSessao},
        {"Bttn_IrPara_Configuracoes_Administrativas", IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Administrativas},
        {"Bttn_IrPara_Configuracoes_Infraestrutura", IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Infraestrutura},
        {"Bttn_IrPara_Configuracoes_Academicas", IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Academicas},
        {"Bttn_IrPara_Configuracoes_Educacionais", IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Educacionais},
        {"Bttn_IrPara_Configuracoes_Eventos", IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Eventos},
        {"Bttn_IrPara_Edicao_Cronograma", IdentificadoresDeBotoesDaRibbon.IrPara_Edicao_Cronograma},
        {"Bttn_Configuracoes_Administrativa_VoltarPara_MenuInicial", IdentificadoresDeBotoesDaRibbon.Configuracoes_Administrativa_VoltarPara_MenuInicial},
        {"Bttn_Configurar_TurnoLetivo", IdentificadoresDeBotoesDaRibbon.Configurar_TurnoLetivo},
        {"Bttn_Configurar_HorariosDeAula", IdentificadoresDeBotoesDaRibbon.Configurar_HorariosDeAula},
        {"Bttn_Configurar_TiposUnidadesCurriculares", IdentificadoresDeBotoesDaRibbon.Configurar_TiposUnidadesCurriculares},
        {"Bttn_Configuracoes_Infraestrutura_VoltarPara_MenuInicial", IdentificadoresDeBotoesDaRibbon.Configuracoes_Infraestrutura_VoltarPara_MenuInicial},
        {"Bttn_Configurar_Blocos", IdentificadoresDeBotoesDaRibbon.Configurar_Blocos},
        {"Bttn_Configurar_Andares", IdentificadoresDeBotoesDaRibbon.Configurar_Andares},
        {"Bttn_Configurar_SalasDeAula", IdentificadoresDeBotoesDaRibbon.Configurar_SalasDeAula},
        {"Bttn_Configuracoes_Academicas_VoltarPara_MenuInicial", IdentificadoresDeBotoesDaRibbon.Configuracoes_Academicas_VoltarPara_MenuInicial},
        {"Bttn_Configurar_Docentes", IdentificadoresDeBotoesDaRibbon.Configurar_Docentes},
        {"Bttn_Configurar_AutorizacoesParaLecionar", IdentificadoresDeBotoesDaRibbon.Configurar_AutorizacoesParaLecionar},
        {"Bttn_Configurar_Atestados", IdentificadoresDeBotoesDaRibbon.Configurar_Atestados},
        {"Bttn_Configuracoes_Educacionais_VoltarPara_MenuInicial", IdentificadoresDeBotoesDaRibbon.Configuracoes_Educacionais_VoltarPara_MenuInicial},
        {"Bttn_Configurar_UnidadeEducacional", IdentificadoresDeBotoesDaRibbon.Configurar_UnidadeEducacional},
        {"Bttn_Configurar_AreaProfissional", IdentificadoresDeBotoesDaRibbon.Configurar_AreaProfissional},
        {"Bttn_Configurar_NomeDoCurso", IdentificadoresDeBotoesDaRibbon.Configurar_NomeDoCurso},
        {"Bttn_Configurar_UnidadesCurriculares", IdentificadoresDeBotoesDaRibbon.Configurar_UnidadesCurriculares},
        {"Bttn_Configurar_CodigoDaTurma", IdentificadoresDeBotoesDaRibbon.Configurar_CodigoDaTurma},
        {"Bttn_Configuracoes_Eventos_VoltarPara_MenuInicial", IdentificadoresDeBotoesDaRibbon.Configuracoes_Eventos_VoltarPara_MenuInicial},
        {"Bttn_Configurar_Feriados", IdentificadoresDeBotoesDaRibbon.Configurar_Feriados},
        {"Bttn_Configurar_Recessos", IdentificadoresDeBotoesDaRibbon.Configurar_Recessos},
        {"Bttn_Configurar_DatasEventuais", IdentificadoresDeBotoesDaRibbon.Configurar_DatasEventuais},
        {"Bttn_Edicao_Cronograma_VoltarPara_MenuInicial", IdentificadoresDeBotoesDaRibbon.Edicao_Cronograma_VoltarPara_MenuInicial},
        {"Bttn_CriarCronograma_ComIa", IdentificadoresDeBotoesDaRibbon.CriarCronograma_ComIa},
        {"Bttn_EditarCronograma_Manualmente", IdentificadoresDeBotoesDaRibbon.EditarCronograma_Manualmente},
        {"Bttn_EditarCronograma_ComIa", IdentificadoresDeBotoesDaRibbon.EditarCronograma_ComIa},
        {"Bttn_Visualizar_Erros", IdentificadoresDeBotoesDaRibbon.Visualizar_Erros},
        {"Bttn_Visualizar_CalendarioAcademico", IdentificadoresDeBotoesDaRibbon.Visualizar_CalendarioAcademico},
        {"Bttn_ExportarEm_PDF", IdentificadoresDeBotoesDaRibbon.ExportarEm_PDF},
        {"Bttn_ExportarEm_XLSX", IdentificadoresDeBotoesDaRibbon.ExportarEm_XLSX}
    }

    Public ReadOnly ExecutorDeAcoesDaRibbon As New Dictionary(Of IdentificadoresDeBotoesDaRibbon, System.Action) From {
        {IdentificadoresDeBotoesDaRibbon.IniciarSessao, Sub() NavegarParaTab("Tb_MenuInicial")},
        {IdentificadoresDeBotoesDaRibbon.FinalizarSessao, Sub() NavegarParaTab("Tb_Logon")},
        {IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Administrativas, Sub() NavegarParaTab("Tb_Configuracoes_Administrativas")},
        {IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Infraestrutura, Sub() NavegarParaTab("Tb_Configuracoes_Infraestrutura")},
        {IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Academicas, Sub() NavegarParaTab("Tb_Configuracoes_Academicas")},
        {IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Educacionais, Sub() NavegarParaTab("Tb_Configuracoes_Educacionais")},
        {IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Eventos, Sub() NavegarParaTab("Tb_Configuracoes_Eventos")},
        {IdentificadoresDeBotoesDaRibbon.IrPara_Edicao_Cronograma, Sub() NavegarParaTab("Tb_Edicao_Cronograma")},
        {IdentificadoresDeBotoesDaRibbon.Configuracoes_Administrativa_VoltarPara_MenuInicial, Sub() NavegarParaTab("Tb_MenuInicial")},
        {IdentificadoresDeBotoesDaRibbon.Configurar_TurnoLetivo, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configurar_HorariosDeAula, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configurar_TiposUnidadesCurriculares, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configuracoes_Infraestrutura_VoltarPara_MenuInicial, Sub() NavegarParaTab("Tb_MenuInicial")},
        {IdentificadoresDeBotoesDaRibbon.Configurar_Blocos, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configurar_Andares, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configurar_SalasDeAula, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configuracoes_Academicas_VoltarPara_MenuInicial, Sub() NavegarParaTab("Tb_MenuInicial")},
        {IdentificadoresDeBotoesDaRibbon.Configurar_Docentes, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configurar_AutorizacoesParaLecionar, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configurar_Atestados, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configuracoes_Educacionais_VoltarPara_MenuInicial, Sub() NavegarParaTab("Tb_MenuInicial")},
        {IdentificadoresDeBotoesDaRibbon.Configurar_UnidadeEducacional, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configurar_AreaProfissional, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configurar_NomeDoCurso, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configurar_UnidadesCurriculares, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configurar_CodigoDaTurma, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configuracoes_Eventos_VoltarPara_MenuInicial, Sub() NavegarParaTab("Tb_MenuInicial")},
        {IdentificadoresDeBotoesDaRibbon.Configurar_Feriados, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configurar_Recessos, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Configurar_DatasEventuais, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Edicao_Cronograma_VoltarPara_MenuInicial, Sub() NavegarParaTab("Tb_MenuInicial")},
        {IdentificadoresDeBotoesDaRibbon.CriarCronograma_ComIa, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.EditarCronograma_Manualmente, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.EditarCronograma_ComIa, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Visualizar_Erros, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.Visualizar_CalendarioAcademico, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.ExportarEm_PDF, Sub() EmBreve()},
        {IdentificadoresDeBotoesDaRibbon.ExportarEm_XLSX, Sub() EmBreve()}
    }

    Public ReadOnly MapaDeIconesDeBotoesDaRibbon As New Dictionary(Of IdentificadoresDeBotoesDaRibbon, Bitmap) From {
        {IdentificadoresDeBotoesDaRibbon.IniciarSessao, My.Resources.icn_IniciarSessao},
        {IdentificadoresDeBotoesDaRibbon.FinalizarSessao, My.Resources.icn_FinalizarSessao},
        {IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Administrativas, My.Resources.icn_ConfiguracoesAdministrativas},
        {IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Infraestrutura, My.Resources.icn_ConfiguracoesdeInfraestrutura},
        {IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Academicas, My.Resources.icn_ConfiguracoesAcadêmicas},
        {IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Educacionais, My.Resources.icn_ConfiguracoesEducacionais},
        {IdentificadoresDeBotoesDaRibbon.IrPara_Configuracoes_Eventos, My.Resources.icn_ConfiguracoesDeEventos},
        {IdentificadoresDeBotoesDaRibbon.IrPara_Edicao_Cronograma, My.Resources.icn_IrParaEditarCronograma},
        {IdentificadoresDeBotoesDaRibbon.Configuracoes_Administrativa_VoltarPara_MenuInicial, My.Resources.icn_VoltarParaMenuInicial},
        {IdentificadoresDeBotoesDaRibbon.Configurar_TurnoLetivo, My.Resources.icn_ConfigurarTurnoLetivo},
        {IdentificadoresDeBotoesDaRibbon.Configurar_HorariosDeAula, My.Resources.icn_ConfigurarHorarisDeAulas},
        {IdentificadoresDeBotoesDaRibbon.Configurar_TiposUnidadesCurriculares, My.Resources.icn_ConfigurarTiposDeUnidadesCurriculares},
        {IdentificadoresDeBotoesDaRibbon.Configuracoes_Infraestrutura_VoltarPara_MenuInicial, My.Resources.icn_VoltarParaMenuInicial},
        {IdentificadoresDeBotoesDaRibbon.Configurar_Blocos, My.Resources.icn_ConfigurarBlocos},
        {IdentificadoresDeBotoesDaRibbon.Configurar_Andares, My.Resources.icn_ConfigurarAndares},
        {IdentificadoresDeBotoesDaRibbon.Configurar_SalasDeAula, My.Resources.icn_ConfigurarSalasDeAula},
        {IdentificadoresDeBotoesDaRibbon.Configuracoes_Academicas_VoltarPara_MenuInicial, My.Resources.icn_VoltarParaMenuInicial},
        {IdentificadoresDeBotoesDaRibbon.Configurar_Docentes, My.Resources.icn_ConfigurarDocentes},
        {IdentificadoresDeBotoesDaRibbon.Configurar_AutorizacoesParaLecionar, My.Resources.icn_ConfigurarAutorizacoesParaLecionar},
        {IdentificadoresDeBotoesDaRibbon.Configurar_Atestados, My.Resources.icn_ConfigurarAtestados},
        {IdentificadoresDeBotoesDaRibbon.Configuracoes_Educacionais_VoltarPara_MenuInicial, My.Resources.icn_VoltarParaMenuInicial},
        {IdentificadoresDeBotoesDaRibbon.Configurar_UnidadeEducacional, My.Resources.icn_ConfigurarUnidadeEducacional},
        {IdentificadoresDeBotoesDaRibbon.Configurar_AreaProfissional, My.Resources.icn_ConfigurarAreaProfissional},
        {IdentificadoresDeBotoesDaRibbon.Configurar_NomeDoCurso, My.Resources.icn_ConfigurarNomeDoCurso},
        {IdentificadoresDeBotoesDaRibbon.Configurar_UnidadesCurriculares, My.Resources.icn_ConfigurarUnidadesCurriculares},
        {IdentificadoresDeBotoesDaRibbon.Configurar_CodigoDaTurma, My.Resources.icn_ConfigurarCodigoDaTurma},
        {IdentificadoresDeBotoesDaRibbon.Configuracoes_Eventos_VoltarPara_MenuInicial, My.Resources.icn_VoltarParaMenuInicial},
        {IdentificadoresDeBotoesDaRibbon.Configurar_Feriados, My.Resources.icn_ConfigurarFeriados},
        {IdentificadoresDeBotoesDaRibbon.Configurar_Recessos, My.Resources.icn_ConfigurarRecessos},
        {IdentificadoresDeBotoesDaRibbon.Configurar_DatasEventuais, My.Resources.icn_ConfigurarDatasEventuais},
        {IdentificadoresDeBotoesDaRibbon.Edicao_Cronograma_VoltarPara_MenuInicial, My.Resources.icn_VoltarParaMenuInicial},
        {IdentificadoresDeBotoesDaRibbon.CriarCronograma_ComIa, My.Resources.icn_CriarCronogramaComIA},
        {IdentificadoresDeBotoesDaRibbon.EditarCronograma_Manualmente, My.Resources.icn_EditarCronogramaManualmente},
        {IdentificadoresDeBotoesDaRibbon.EditarCronograma_ComIa, My.Resources.icn_EditarCronogramaComIA},
        {IdentificadoresDeBotoesDaRibbon.Visualizar_Erros, My.Resources.icn_VisualizarErros},
        {IdentificadoresDeBotoesDaRibbon.Visualizar_CalendarioAcademico, My.Resources.icn_VisualizarCalendarioAcademico},
        {IdentificadoresDeBotoesDaRibbon.ExportarEm_PDF, My.Resources.icn_ExportarEmPDF},
        {IdentificadoresDeBotoesDaRibbon.ExportarEm_XLSX, My.Resources.icn_ExportarEmXLSX}
    }


    Public ReadOnly RegrasDeHabilitacaoDeBotoesDaRibbon As New Dictionary(Of String, Func(Of Boolean)) From {
        {"Bttn_IniciarSessao", Function() True},
        {"Bttn_FinalizarSessao", Function() True},
        {"Bttn_IrPara_Configuracoes_Administrativas", Function() True},
        {"Bttn_IrPara_Configuracoes_Infraestrutura", Function() VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__TIPOS_DE_UNIDADES_CURRICULARES") > 0},
        {"Bttn_IrPara_Configuracoes_Academicas", Function() VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__SALAS_DE_AULA") > 0},
        {"Bttn_IrPara_Configuracoes_Educacionais", Function() VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__DOCENTES") > 0},
        {"Bttn_IrPara_Configuracoes_Eventos", Function() Not String.IsNullOrEmpty(ObterTexto(PL_CONFIGURACOES, "RNG__CONFIGURACAO__DADOS_TURMA__CODIGO_DA_TURMA"))},
        {"Bttn_IrPara_Edicao_Cronograma", Function()
                                              Return VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__FERIADOS") > 0 AndAlso
                                              VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__RECESSOS") > 0
                                          End Function},
        {"Bttn_Configuracoes_Administrativa_VoltarPara_MenuInicial", Function() True},
        {"Bttn_Configurar_TurnoLetivo", Function() True},
        {"Bttn_Configurar_HorariosDeAula", Function() Not String.IsNullOrEmpty(ObterTexto(PL_CONFIGURACOES, "RNG__CONFIGURACAO__DADOS_TURMA__TURNO_LETIVO"))},
        {"Bttn_Configurar_TiposUnidadesCurriculares", Function() VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__HORARIOS") > 0},
        {"Bttn_Configuracoes_Infraestrutura_VoltarPara_MenuInicial", Function() True},
        {"Bttn_Configurar_Blocos", Function() True},
        {"Bttn_Configurar_Andares", Function() VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__BLOCOS") > 0},
        {"Bttn_Configurar_SalasDeAula", Function() VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__ANDARES") > 0},
        {"Bttn_Configuracoes_Academicas_VoltarPara_MenuInicial", Function() True},
        {"Bttn_Configurar_Docentes", Function() True},
        {"Bttn_Configurar_AutorizacoesParaLecionar", Function() VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__DOCENTES") > 0},
        {"Bttn_Configurar_Atestados", Function() VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__DOCENTES") > 0},
        {"Bttn_Configuracoes_Educacionais_VoltarPara_MenuInicial", Function() True},
        {"Bttn_Configurar_UnidadeEducacional", Function() True},
        {"Bttn_Configurar_AreaProfissional", Function() Not String.IsNullOrEmpty(ObterTexto(PL_CONFIGURACOES, "RNG__CONFIGURACAO__DADOS_TURMA__UNIDADE_EDUCACIONAL"))},
        {"Bttn_Configurar_NomeDoCurso", Function() Not String.IsNullOrEmpty(ObterTexto(PL_CONFIGURACOES, "RNG__CONFIGURACAO__DADOS_TURMA__AREA_PROFISSIONAL"))},
        {"Bttn_Configurar_UnidadesCurriculares", Function() Not String.IsNullOrEmpty(ObterTexto(PL_CONFIGURACOES, "RNG__CONFIGURACAO__DADOS_TURMA__NOME_CURSO"))},
        {"Bttn_Configurar_CodigoDaTurma", Function() VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__UNIDADES_CURRICULARES") > 0},
        {"Bttn_Configuracoes_Eventos_VoltarPara_MenuInicial", Function() True},
        {"Bttn_Configurar_Feriados", Function() True},
        {"Bttn_Configurar_Recessos", Function() VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__FERIADOS") > 0},
        {"Bttn_Configurar_DatasEventuais", Function()
                                               Return VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__FERIADOS") > 0 AndAlso
                                               VerificarQtd(PL_CONFIGURACOES, "RNG__CONFIGURACAO__QUANTIDADES_REGISTROS__RECESSOS") > 0
                                           End Function},
        {"Bttn_Edicao_Cronograma_VoltarPara_MenuInicial", Function() True},
        {"Bttn_CriarCronograma_ComIa", Function() True},
        {"Bttn_EditarCronograma_Manualmente", Function() True},
        {"Bttn_EditarCronograma_ComIa", Function() True},
        {"Bttn_Visualizar_Erros", Function() True},
        {"Bttn_Visualizar_CalendarioAcademico", Function() True},
        {"Bttn_ExportarEm_PDF", Function() True},
        {"Bttn_ExportarEm_XLSX", Function() True}
    }


#End Region

End Module
