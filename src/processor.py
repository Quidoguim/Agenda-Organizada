import pandas as pd


def contar_atividades_por_pessoa_e_dia(df_agenda):
    """
    Conta a quantidade de atividades e atos por pessoa em cada dia.
    Considera a coluna "Responsável" como nome da pessoa.
    """
    df_agenda['Data'] = pd.to_datetime(df_agenda['Data'])

    # Contagem de atividades por Responsável e Data
    atividades = df_agenda.groupby(
        ['Responsável', 'Data']).size().reset_index(name='Qtd Atividades')

    # Contagem de atos por Responsável e Data, considerando "Cobrar ato" na coluna 'Etiquetas'
    df_agenda['Contar Ato'] = df_agenda['Etiquetas'].astype(
        str).apply(lambda x: 'Cobrar ato' in x)
    atos = df_agenda[df_agenda['Contar Ato']].groupby(
        ['Responsável', 'Data']).size().reset_index(name='Qtd Atos')

    # Mesclar as contagens de atividades e atos
    resultado_completo = pd.merge(
        atividades, atos, on=['Responsável', 'Data'], how='left')
    resultado_completo['Qtd Atos'] = resultado_completo['Qtd Atos'].fillna(
        0).astype(int)

    # Organizar por responsável
    resultado = {}
    for nome, grupo in resultado_completo.groupby('Responsável'):
        resultado[nome] = grupo.drop(
            columns='Responsável').reset_index(drop=True)

    return resultado
