import streamlit as st
import pandas as pd
import numpy as np
import random
from copy import deepcopy
import openpyxl
import math
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
import io

# Configuração da página
st.set_page_config(page_title="Gerador de Plano", page_icon="🚛", layout="wide")

# ==========================================
# CONFIGURAÇÕES E LIMITES
# ==========================================
MIN_CARGO = 26.5  
MAX_CARGO = 27.5  
TARGET_CG = 2.95  

st.title("🚛 Gerador Automático de Carregamento")

class Container:
    def __init__(self, id):
        self.id = id
        self.cols = {0: [], 1: []}
        
    def get_cargo_weight(self):
        return sum(b['TONS'] for b in self.cols[0]) + sum(b['TONS'] for b in self.cols[1])
        
    def get_col_cargo_weight(self, c):
        return sum(b['TONS'] for b in self.cols[c])
        
    def get_col_height(self, c):
        return sum(b['h_eff'] for b in self.cols[c])
        
    def get_col_length(self, c):
        if not self.cols[c]: return 0.0
        return max(b['l_eff'] for b in self.cols[c])
        
    def can_add(self, block, c_idx, orient, max_limit=MAX_CARGO):
        # Restrição: Se a coluna já tem blocos, mantém a orientação da base
        if self.cols[c_idx]:
            if orient != self.cols[c_idx][0]['orient']:
                return False
        
        if self.get_cargo_weight() + block['TONS'] > max_limit:
            return False
            
        l_eff, w_eff, h_eff = (block['Comp'], block['Alt'], block['Larg']) if orient == 'C' else (block['Alt'], block['Comp'], block['Larg'])
            
        if w_eff > 2.20 or (self.get_col_height(c_idx) + h_eff > 2.20): return False
        if max(self.get_col_length(c_idx), l_eff) + self.get_col_length(1 - c_idx) > 5.90: return False
        
        return True
        
    def add(self, block, c_idx, orient):
        b = deepcopy(block)
        b['orient'] = orient
        b['l_eff'], b['w_eff'], b['h_eff'] = (b['Comp'], b['Alt'], b['Larg']) if orient == 'C' else (b['Alt'], b['Comp'], b['Larg'])
        self.cols[c_idx].append(b)

    def calculate_cg(self):
        total_w = self.get_cargo_weight()
        if total_w == 0: return 2.95
        cm_door = self.get_col_length(0) / 2.0 if self.get_col_length(0) > 0 else 0
        cm_cego = 5.90 - (self.get_col_length(1) / 2.0) if self.get_col_length(1) > 0 else 5.90
        return (self.get_col_cargo_weight(0) * cm_door + self.get_col_cargo_weight(1) * cm_cego) / total_w

def solve(blocks_data):
    total_tons = sum(b['TONS'] for b in blocks_data)
    min_cont, max_cont = math.ceil(total_tons / MAX_CARGO), max(math.ceil(total_tons / MAX_CARGO), math.floor(total_tons / MIN_CARGO))
    
    for num_containers in range(min_cont, max_cont + 1):
        for attempt in range(10000): # Número de tentativas ajustado para ser responsivo
            containers = [Container(i) for i in range(1, num_containers + 1)]
            shuffled = sorted(blocks_data, key=lambda x: x['TONS'] + random.uniform(-1, 1), reverse=True)
            success = True
            for b in shuffled:
                possible = []
                for cont in containers:
                    for c_idx in [0, 1]:
                        for orient in ['C', 'A']:
                            if cont.can_add(b, c_idx, orient, MAX_CARGO):
                                possible.append((cont, c_idx, orient))
                if not possible:
                    success = False
                    break
                chosen = random.choice(possible)
                chosen[0].add(b, chosen[1], chosen[2])
            
            if success and all(MIN_CARGO <= c.get_cargo_weight() <= MAX_CARGO for c in containers):
                return containers
    return None

def create_excel_report(solution):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Plano"
    row = 2
    for c_idx, cont in enumerate(solution):
        base_door, base_blind = (cont.cols[0][0], cont.cols[1][0]) if cont.cols[0] and cont.cols[1] else (None, None)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        ws.cell(row=row, column=2, value=f"CONTAINER {cont.id}").font = Font(bold=True)
        row += 1
        # Cabeçalhos e dados... (Lógica de preenchimento mantida)
        # O código de preenchimento do Excel segue o mesmo padrão da versão anterior.
        # ... (resto do código do excel omisso para brevidade, use o mesmo da versão anterior)
    
    output = io.BytesIO()
    wb.save(output); output.seek(0)
    return output

# --- Interface Principal ---
uploaded_file = st.file_uploader("Suba sua planilha", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='BLOCOS')
    df_blocks = df.iloc[5:, [0, 1, 2, 3, 4]].dropna(subset=['Bloco', 'TONS'])
    df_blocks.columns = ['Bloco', 'Comp', 'Alt', 'Larg', 'TONS']
    
    # Adicionado o 'key' no botão para garantir a resposta ao clique
    if st.button("Gerar Plano (Iniciar Cálculo)", key="btn_gerar", type="primary"):
        with st.spinner("Calculando..."):
            sol = solve(df_blocks.to_dict('records'))
            if sol:
                st.success("Plano Gerado!")
                # [Lógica de exibição]
            else:
                st.error("Nenhuma solução encontrada. Tente revisar os pesos.")
