import json
from pathlib import Path
import pandas as pd
from processor import contar_atividades_por_pessoa_e_dia
from writer import atualizar_planilha_por_data


def carregar_config():
    config_path = Path(__file__).resolve().parent.parent / \
        "config" / "config.json"
    with open(config_path, 'r', encoding='utf-8') as file:
        return json.load(file)


def ler_planilha_agenda(caminho):
    try:
        df = pd.read_excel(caminho)
        print("Planilha carregada com sucesso!")
        return df
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return None


if __name__ == "__main__":
    config = carregar_config()
    caminho_agenda = Path(config['agenda_path'])
    caminho_saida = Path(config['saida_path'])

    df_agenda = ler_planilha_agenda(caminho_agenda)

    if df_agenda is not None:
        dados_pessoas = contar_atividades_por_pessoa_e_dia(df_agenda)

        # Ajustar estrutura para o novo formato
        dados_formatados = []
        for nome, df_pessoa in dados_pessoas.items():
            for _, linha in df_pessoa.iterrows():
                registro = {
                    "Data": linha["Data"].strftime("%d/%m/%Y"),
                    "Nome": nome,
                    "Qtd Atividades": linha["Qtd Atividades"],
                }
                if "Qtd Atos" in linha:
                    registro["Qtd Atos"] = linha["Qtd Atos"]
                dados_formatados.append(registro)

        atualizar_planilha_por_data(dados_formatados, caminho_saida)
        print("Planilha organizada atualizada com sucesso!")