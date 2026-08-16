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
st.set_page_config(page_title="Gerador de Plano de Carregamento", page_icon="🚛", layout="wide")

# ==========================================
# CONFIGURAÇÕES E LIMITES
# ==========================================
MIN_CARGO = 26.5  
MAX_CARGO = 27.5  
TARGET_CG = 2.95  

st.title("🚛 Gerador Automático de Carregamento (Padrão Diretoria)")
st.write("Constraint aplicada: Se o bloco da base define a orientação (A ou C), todos os blocos acima seguem o mesmo padrão.")

class Container:
    def __init__(self, id):
        self.id = id
        self.cols = {0: [], 1: []}  # 0: Porta, 1: Cego
        
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
        # --- RESTRICAO: Toda a coluna deve ter a mesma orientação da base ---
        if self.cols[c_idx]:
            if orient != self.cols[c_idx][0]['orient']:
                return False
        
        # Limite de carga
        if self.get_cargo_weight() + block['TONS'] > max_limit:
            return False
            
        # Cálculo de dimensões efetivas
        if orient == 'C':
            l_eff, w_eff, h_eff = block['Comp'], block['Alt'], block['Larg']
        else: 
            l_eff, w_eff, h_eff = block['Alt'], block['Comp'], block['Larg']
            
        if w_eff > 2.20: return False
        if self.get_col_height(c_idx) + h_eff > 2.20: return False
        
        # Limite de comprimento total (5.90m)
        new_col_len = max(self.get_col_length(c_idx), l_eff)
        other_col_len = self.get_col_length(1 - c_idx)
        if new_col_len + other_col_len > 5.90: return False
        
        return True
        
    def add(self, block, c_idx, orient):
        b = deepcopy(block)
        b['orient'] = orient
        if orient == 'C':
            b['l_eff'], b['w_eff'], b['h_eff'] = b['Comp'], b['Alt'], b['Larg']
        else:
            b['l_eff'], b['w_eff'], b['h_eff'] = b['Alt'], b['Comp'], b['Larg']
        self.cols[c_idx].append(b)

    def calculate_cg(self):
        door_w = self.get_col_cargo_weight(0)
        cego_w = self.get_col_cargo_weight(1)
        total_w = door_w + cego_w
        if total_w == 0: return 2.95
        
        cm_door = self.get_col_length(0) / 2.0 if self.get_col_length(0) > 0 else 0
        cm_cego = 5.90 - (self.get_col_length(1) / 2.0) if self.get_col_length(1) > 0 else 5.90
        
        cg = (door_w * cm_door + cego_w * cm_cego) / total_w
        return cg

def solve(blocks_data):
    total_tons = sum(b['TONS'] for b in blocks_data)
    min_cont = math.ceil(total_tons / MAX_CARGO)
    max_cont = max(min_cont, math.floor(total_tons / MIN_CARGO))
    
    best_solution = None
    best_score_val = float('inf')
    
    tolerance_steps = [(MIN_CARGO, MAX_CARGO, 50000), (25.5, 28.0, 70000)]
    
    for min_l, max_l, attempts in tolerance_steps:
        for num_containers in range(min_cont, max_cont + 1):
            target_weight = total_tons / num_containers
            
            for attempt in range(attempts):
                containers = [Container(i) for i in range(1, num_containers + 1)]
                shuffled = sorted(blocks_data, key=lambda x: x['TONS'] + random.uniform(-2.0, 2.0), reverse=True)
                
                success = True
                for b in shuffled:
                    valid_moves = []
                    for cont in containers:
                        for c_idx in [0, 1]:
                            for orient in ['C', 'A']:
                                if cont.can_add(b, c_idx, orient, max_l):
                                    # Simulação para penalidade
                                    door_w = cont.get_col_cargo_weight(0) + (b['TONS'] if c_idx == 0 else 0)
                                    cego_w = cont.get_col_cargo_weight(1) + (b['TONS'] if c_idx == 1 else 0)
                                    door_l = max(cont.get_col_length(0), (b['Comp'] if orient=='C' else b['Alt']) if c_idx == 0 else 0)
                                    cego_l = max(cont.get_col_length(1), (b['Comp'] if orient=='C' else b['Alt']) if c_idx == 1 else 0)
                                    
                                    cm_door = door_l / 2.0 if door_l > 0 else 0
                                    cm_cego = 5.90 - (cego_l / 2.0) if cego_l > 0 else 5.90
                                    
                                    simulated_total_w = door_w + cego_w
                                    simulated_cg = (door_w * cm_door + cego_w * cm_cego) / simulated_total_w
                                    
                                    cg_penalty = abs(simulated_cg - TARGET_CG) * 120
                                    target_dist = abs(simulated_total_w - target_weight) * 30
                                    pattern_penalty = (max(0, 6.0 - door_w) + max(0, door_w - 13.0) + max(0, 13.5 - cego_w) + max(0, cego_w - 20.5)) * 6
                                    
                                    score = cg_penalty + target_dist + pattern_penalty
                                    valid_moves.append((score, cont, c_idx, orient))
                    
                    if not valid_moves:
                        success = False
                        break
                    else:
                        valid_moves.sort(key=lambda x: x[0])
                        chosen = random.choice(valid_moves[: min(3, len(valid_moves))])
                        chosen[1].add(b, chosen[2], chosen[3])
                        
                if success:
                    is_valid = all(min_l <= c.get_cargo_weight() <= max_l for c in containers)
                    if is_valid:
                        total_score = sum(abs(c.calculate_cg() - TARGET_CG) for c in containers) * 50
                        if total_score < best_score_val:
                            best_score_val = total_score
                            best_solution = deepcopy(containers)
                            
        if best_solution is not None:
            return best_solution, True
            
    return None, False

def create_excel_report(solution):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Plano de Carregamento"
    colors = ['FFFF00', '00B0F0', '00FF00', 'FF9900', 'CCC0DA', 'FF0000', 'E2EFDA', 'FCE4D6', 'D9E1F2']
    thick_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    row_offset = 2
    for c_idx, cont in enumerate(solution):
        # A base real da coluna (primeiro bloco inserido)
        base_door = cont.cols[0][0] if len(cont.cols[0]) > 0 else None
        base_blind = cont.cols[1][0] if len(cont.cols[1]) > 0 else None
        
        # Ordenação para exibição
        door_col = sorted(cont.cols[0], key=lambda x: x['l_eff'], reverse=True)
        blind_col = sorted(cont.cols[1], key=lambda x: x['l_eff'], reverse=True)
        
        # Cabeçalho
        ws.merge_cells(start_row=row_offset, start_column=2, end_row=row_offset, end_column=4)
        cell = ws.cell(row=row_offset, column=2, value=f"CONTAINER {cont.id} (CG: {cont.calculate_cg():.2f}m)")
        cell.font = Font(bold=True, size=14); cell.alignment = Alignment(horizontal='center')
        
        row_offset += 1
        ws.cell(row=row_offset, column=2, value="Lado Porta").alignment = Alignment(horizontal='center')
        ws.cell(row=row_offset, column=3, value="Lado Cego").alignment = Alignment(horizontal='center')
        row_offset += 1
        
        # Orientação da base real (A ou C)
        ws.cell(row=row_offset, column=2, value=base_door['orient'] if base_door else "").alignment = Alignment(horizontal='center')
        ws.cell(row=row_offset, column=3, value=base_blind['orient'] if base_blind else "").alignment = Alignment(horizontal='center')
        row_offset += 1
        
        max_blocks = max(len(door_col), len(blind_col))
        fill = PatternFill(start_color=colors[c_idx % len(colors)], end_color=colors[c_idx % len(colors)], fill_type="solid")
        
        for i in range(max_blocks - 1, -1, -1):
            if i < len(door_col):
                c = ws.cell(row=row_offset, column=2, value=door_col[i]['Bloco'])
                c.fill = fill; c.border = thick_border; c.alignment = Alignment(horizontal='center')
            if i < len(blind_col):
                c = ws.cell(row=row_offset, column=3, value=blind_col[i]['Bloco'])
                c.fill = fill; c.border = thick_border; c.alignment = Alignment(horizontal='center')
            row_offset += 1
            
        # Comprimentos calculados com base na orientação da base real
        door_len = base_door['l_eff'] if base_door else 0
        blind_len = base_blind['l_eff'] if base_blind else 0
        ws.cell(row=row_offset, column=2, value=round(door_len, 2)).alignment = Alignment(horizontal='center')
        ws.cell(row=row_offset, column=3, value=round(blind_len, 2)).alignment = Alignment(horizontal='center')
        ws.cell(row=row_offset, column=4, value=round(door_len + blind_len, 2)).alignment = Alignment(horizontal='center')
        row_offset += 1
        
        # Pesos
        ws.cell(row=row_offset, column=2, value=round(cont.get_col_cargo_weight(0), 3)).alignment = Alignment(horizontal='center')
        ws.cell(row=row_offset, column=3, value=round(cont.get_col_cargo_weight(1), 3)).alignment = Alignment(horizontal='center')
        ws.cell(row=row_offset, column=4, value=round(cont.get_cargo_weight(), 3)).alignment = Alignment(horizontal='center')
        row_offset += 3
        
    output = io.BytesIO()
    wb.save(output); output.seek(0)
    return output

# --- Interface Principal ---
uploaded_file = st.file_uploader("Selecione sua planilha de blocos (Excel)", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='BLOCOS')
    df_blocks = df.iloc[5:, [0, 1, 2, 3, 4]].copy() 
    df_blocks.columns = ['Bloco', 'Comp', 'Alt', 'Larg', 'TONS']
    blocks_data = df_blocks.dropna(subset=['Bloco', 'TONS']).to_dict('records')
    
    if st.button("Gerar Plano Padrão Diretoria", type="primary"):
        sol, _ = solve(blocks_data)
        if sol:
            st.subheader("Resumo do Carregamento:")
            for c in sol:
                st.write(f"**Container {c.id}** | Carga: {c.get_cargo_weight():.2f} t | CG: {c.calculate_cg():.2f} m")
            st.download_button("📥 Baixar Excel", data=create_excel_report(sol), file_name="Plano_Carregamento.xlsx")
        else:
            st.error("Não foi possível encontrar uma solução válida.")
