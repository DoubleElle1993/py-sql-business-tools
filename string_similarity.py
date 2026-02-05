import openpyxl
import os
import re
from rapidfuzz import fuzz
import pandas as pd

# Paths
base_dir = os.getcwd()
file_path = os.path.join(base_dir, '')
out_path = os.path.join(os.path.dirname(file_path), '')


def normalize_text(s: str) -> str:
    """Funzione di normalizzazione delle string"""
    if pd.isna(s):
        return ''
    s = str(s).lower().strip()
    s = s.replace("-", " ").replace("_", " ")
    return s


def similarity_ratio(a_n: str, b_n: str) -> float:
    """Funzione di similarità con valori da 0 a 1"""
    if not a_n or not b_n:
        return 0.0
    return fuzz.ratio(a_n, b_n) / 100.0


def similarity_unordered(a_n: str, b_n: str) -> float:
    """Funzione di similarità che ignora l'ordine delle parole"""
    if not a_n or not b_n:
        return 0.0
    return fuzz.token_set_ratio(a_n, b_n) / 100.0


def main(path):
    """Funzione principale di analisi testuale sul file"""
    threshold = 0.6
    file_trader = pd.read_excel(file_path, sheet_name='Trader', engine='openpyxl')
    file_mafu = pd.read_excel(file_path, sheet_name='Mafu', engine='openpyxl')

    file_mafu['collection_norma'] = file_mafu['Collection'].apply(normalize_text)
    # file_mafu['Collection_norm'] = file_mafu['Collection'].apply(normalize_text)
    file_trader['collection_norma'] = file_trader['collection'].apply(normalize_text)

    #Lista valori di Trader
    mafu_list = file_mafu['collection_norma'].tolist()
    best_ids = []
    best_scores = []

    for flowId_norm in file_trader['collection_norma']:
        if not flowId_norm:
            best_ids.append('')
            best_scores.append(0.0)
            continue

        best_score = 0.0
        best_id = ''

        for mafu_norm in mafu_list:
            if not mafu_norm:
                continue
            score_simil = similarity_unordered(flowId_norm, mafu_norm)
            if score_simil > best_score:
                best_score = score_simil
                best_id = mafu_norm

        best_scores.append(best_score)
        best_ids.append(best_id)

    file_trader['best_mafu'] = best_ids
    file_trader['best_score'] = best_scores
    file_trader['match'] = file_trader['best_score'] >= threshold

    file_trader.to_excel(out_path, index=False)

if __name__ == '__main__':

    main(file_path)
