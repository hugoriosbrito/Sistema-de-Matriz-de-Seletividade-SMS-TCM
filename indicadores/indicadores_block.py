from indicadores._indicadores import Indicador
import customtkinter as ctk

def indicadores_block(frame,sheet):
    """"
    Função para juntar o bloco de indicadores, criando a partir da classe Indicador de _indicadores
    """
    fonte_colunas = ctk.CTkFont(family='Arial', size=15, weight='bold')

    coluna_risco = ctk.CTkLabel(master=frame, text= "RISCO", font=fonte_colunas, text_color='white', corner_radius=20)
    coluna_risco.grid(padx=10,pady=5,row=0,column=0)

    coluna_relevancia = ctk.CTkLabel(master=frame, text= "RELEVÂNCIA", font=fonte_colunas, text_color='white', corner_radius=20)
    coluna_relevancia.grid(padx=10,pady=5,row=0,column=1)

    coluna_materialidade = ctk.CTkLabel(master=frame, text= "MATERIALIDADE", font=fonte_colunas, text_color='white', corner_radius=20)
    coluna_materialidade.grid(padx=10,pady=5,row=0,column=2)

    coluna_oportunidade = ctk.CTkLabel(master=frame, text= "OPORTUNIDADE", font=fonte_colunas, text_color='white', corner_radius=20)
    coluna_oportunidade.grid(padx=0,pady=5,row=0,column=3)

    """
    Indicadores:
    """

    #Tipo Risco
    i_historico_parecer_previo_ultimos_3_anos = Indicador(tipo="risco",
                                                          nome="HISTÓRICO PARECER PRÉVIO",
                                                          descricao='(ÚLTIMOS 3 ANOS)',
                                                          celula_xlsx='F11',
                                                          sheet=sheet,
                                                          frame=frame)


    i_QTDE_DE_DEBITO_MULTAS = Indicador(tipo="risco",
                                        nome="QTDE DE DÉBITO/MULTAS",
                                        descricao='indicadorteste',
                                        celula_xlsx='F13',
                                        sheet=sheet,
                                        frame=frame)


    i_INDICE_DE_TRANSPARENCIA_PUBLICA = Indicador(tipo="risco",
                                                  nome="ÍNDICE DE TRANSPARÊNCIA PÚBLICA",
                                                  descricao='indicadorteste',
                                                  celula_xlsx='F15',
                                                  sheet= sheet,
                                                  frame=frame)

    i_PERFIL_DE_CONTRATACAO_DO_ENTE = Indicador(tipo="risco",
                                                nome="PERFIL DE CONTRATAÇÃO DO ENTE",
                                                descricao='indicadorteste',
                                                celula_xlsx='F16',
                                                sheet=sheet,
                                                frame=frame)
    i_QTDE_DE_DENUNCIAS_E_REPRESENTACOES_ULTIMOS_5_ANOS = Indicador(tipo="risco",
                                                                    nome="QTDE DE DENÚNCIAS E REPRESENTAÇÕES",
                                                                    descricao='(ÚLTIMOS 5 ANOS)',
                                                                    celula_xlsx='F21',
                                                                    sheet=sheet,
                                                                    frame=frame)
    i_QTDE_DE_TOC_ULTIMOS_5_ANOS = Indicador(tipo="risco",
                                            nome="QTDE DE TOC",
                                            descricao='(ÚLTIMOS 5 ANOS)',
                                            celula_xlsx='F22',
                                            sheet=sheet,
                                            frame=frame)

    i_QTDE_DE_TCE_ULTIMOS_5_ANOS = Indicador(tipo="risco",
                                             nome="QTDE DE TCE",
                                             descricao='(ÚLTIMOS 5 ANOS)',
                                             celula_xlsx='F23',
                                             sheet=sheet,
                                             frame=frame)

    i_QTDE_DE_AUDITORIAS_ULTIMOS_5_ANOS = Indicador(tipo="risco",
                                             nome="QTDE DE AUDITORIAS",
                                             descricao='(ÚLTIMOS 5 ANOS',
                                             celula_xlsx='F24',
                                             sheet=sheet,
                                             frame=frame)

    i_QTDE_DE_MEDIDAS_CAUTELARES_ULTIMOS_5_ANOS = Indicador(tipo="risco",
                                                    nome="QTDE DE  MEDIDAS CAUTELARES",
                                                    descricao='(ÚLTIMOS 5 ANOS)',
                                                    celula_xlsx='F25',
                                                    sheet=sheet,
                                                    frame=frame)


    #Tipo Relevância

    i_POPULACAO_MUNICIPIO= Indicador(tipo="relevancia",
                                     nome="POPULAÇÃO MUNICÍPIO",
                                     descricao='indicadorrelevanciateste',
                                     celula_xlsx='F17',
                                     sheet=sheet,
                                     frame=frame)


    i_IDH = Indicador(tipo="relevancia",
                      nome="IDH",
                      descricao='indicadorrelevanciateste',
                      celula_xlsx='F18',
                      sheet= sheet,
                      frame=frame)

    i_IEGM = Indicador(tipo="relevancia",
                       nome="IEGM",
                       descricao='indicadorrelevanciateste',
                       celula_xlsx='F19',
                       sheet=sheet,
                       frame=frame)

    i_IDTRU_DL = Indicador(tipo="relevancia",
                           nome="IDTRU-DL",
                           descricao='indicadorrelevanciateste',
                           celula_xlsx='F20',
                           sheet=sheet,
                           frame=frame)



    #Tipo Materialidade

    i_VALOR_DE_DEBITO_E_MULTAS = Indicador(tipo="materialidade",
                                           nome="VALOR DE DÉBITO E MULTAS",
                                           descricao='indicador teste materialidade',
                                           celula_xlsx='F14',
                                           sheet=sheet,
                                           frame=frame)


    #Tipo Oportunidade

    i_DATA_ULTIMA_AUDITORIA_3DCE = Indicador(tipo="oportunidade",
                                           nome="DATA ÚLTIMA AUDITORIA (3DCE)",
                                           descricao='indicador teste oportunidade',
                                           celula_xlsx='F12',
                                           sheet=sheet,
                                           frame=frame)








