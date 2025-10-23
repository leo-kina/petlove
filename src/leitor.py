from pathlib import Path
import logging

def ler_txt(caminho, separador=","):
    caminho = Path(caminho)
    if not caminho.exists():
        logging.error(f"Arquivo não encontrado: {caminho}")
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    
    dados = []
    with caminho.open("r", encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if linha:
                dados.append(linha.split(separador))
    logging.info(f"{len(dados)} linhas lidas de {caminho}")
    return dados
