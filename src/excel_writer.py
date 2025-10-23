from openpyxl import Workbook
from pathlib import Path
import logging

def criar_planilha(dados, caminho_saida):
    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"

    ws.append(["Nome", "Idade", "Cidade"])
    for linha in dados:
        ws.append(linha)

    caminho_saida = Path(caminho_saida)
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(caminho_saida)
    logging.info(f"Planilha salva em {caminho_saida}")
