import streamlit as st
import pandas as pd
import numpy as np
import random
from copy import deepcopy
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
import io

st.set_page_config(page_title="Gerador de Plano de Carregamento", page_icon="🚛")

st.title("🚛 Gerador Automático de Carregamento")
st.write("Faça o upload da planilha com os blocos e o sistema calculará a melhor distribuição nos containers priorizando o equilíbrio de içamento.")

class Container:
    def __init__(self, id):
        self.id = id
        self.cols = {0: [], 1: []}  # 0: Lado da Porta (Traseira), 1: Lado Cego (Dianteira)
        
    def get_cargo_weight(self):
        return sum(b['TONS'] for b in self.cols[0]) + sum(b['TONS'] for b in self.cols[1])
        
    def get_col_cargo_weight(self, c):
        return sum(b['TONS'] for b in self.cols[c])
        
    def get_total_col_weight(self, c):
        # c=0: Lado da porta (~1.19t estrutural) | c=1: Lado cego (~1.01t estrutural)
        structural_weight = 1.19 if c == 0 else 1.01
        return self.get_col_cargo_weight(c) + structural_weight
        
    def get_total_weight(self):
        return self.get_cargo_weight() + 2.20
        
    def get_col_height(self, c):
        return sum(b['h_eff'] for b in self.cols[c])
        
    def get_col_length(self, c):
        if not self.cols[c]: return 0.0
        return max(b['l_eff'] for b in self.cols[c])
        
    def can_add(self, block, c_idx, orient):
        if self.get_cargo_weight() + block['TONS'] > 25.3:
            return False
            
        if orient == 'C':
            l_eff, w_eff, h_eff = block['Comp'], block['Alt'], block['Larg']
        else: 
            l_eff, w_eff, h_eff = block['Alt'], block['Comp'], block['Larg']
            
        if w_eff > 2.20: return False
        if self.get_col_height(c_idx) + h_eff > 2.20: return False
        
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

def solve(blocks_data):
    best_solution = None
    best_score = float('inf')
    
    for attempt in range(10000):
        containers = [Container(i) for i in range(1, 7)]
        shuffled = sorted(blocks_data, key=lambda x: x['TONS'] + random.uniform(-2, 2), reverse=True)
        
        success = True
        for b in shuffled:
            valid_moves = []
            for cont in containers:
                for c_idx in [0, 1]:
                    for orient in ['C', 'A']:
                        if cont.can_add(b, c_idx, orient):
                            t0 = cont.get_total_col_weight(0) + (b['TONS'] if c_idx == 0 else 0)
                            t1 = cont.get_total_col_weight(1) + (b['TONS'] if c_idx == 1 else 0)
                            
                            balance_penalty = abs(t0 - t1) * 15
                            weight_penalty = cont.get_cargo_weight()
                            new_len = max(cont.get_col_length(c_idx), (b['Comp'] if orient=='C' else b['Alt']))
                            len_penalty = new_len - cont.get_col_length(c_idx)
                            
                            score = weight_penalty * 2 + balance_penalty + len_penalty * 5
                            valid_moves.append((score, cont, c_idx, orient))
            
            if not valid_moves:
                success = False
                break
                
            valid_moves.sort(key=lambda x: x[0])
            chosen = random.choice(valid_moves[:3])
            chosen[1].add(b, chosen[2], chosen[3])
            
        if success:
            total_imbalance = sum(abs(c.get_total_col_weight(0) - c.get_total_col_weight(1)) for c in containers)
            if total_imbalance < best_score:
                best_score = total_imbalance
                best_solution = deepcopy(containers)
                
    return best_solution

def create_excel_report(solution):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Plano de Carregamento"
    colors = ['FFFF00', '00B0F0', '00FF00', 'FF9900', 'CCC0DA', 'FF0000']
    thick_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    row_offset = 2
    for c_idx, cont in enumerate(solution):
        door_col = sorted(cont.cols[0], key=lambda x: x['l_eff'], reverse=True)
        blind_col = sorted(cont.cols[1], key=lambda x: x['l_eff'], reverse=True)
        
        ws.merge_cells(start_row=row_offset, start_column=2, end_row=row_offset, end_column=4)
        cell = ws.cell(row=row_offset, column=2, value=f"CONTAINER {cont.id}")
        cell.font = Font(bold=True, size=14)
        cell.alignment = Alignment(horizontal='center')
        
        row_offset += 1
        ws.cell(row=row_offset, column=2, value="Lado Porta").alignment = Alignment(horizontal='center')
        ws.cell(row=row_offset, column=3, value="Lado Cego").alignment = Alignment(horizontal='center')
        row_offset += 1
        
        max_blocks = max(len(door_col), len(blind_col))
        
        door_orient = door_col[-1]['orient'] if door_col else ""
        blind_orient = blind_col[-1]['orient'] if blind_col else ""
        
        ws.cell(row=row_offset, column=2, value=door_orient).alignment = Alignment(horizontal='center')
        ws.cell(row=row_offset, column=3, value=blind_orient).alignment = Alignment(horizontal='center')
        row_offset += 1
        
        fill = PatternFill(start_color=colors[c_idx % len(colors)], end_color=colors[c_idx % len(colors)], fill_type="solid")
        
        for i in range(max_blocks - 1, -1, -1):
            if i < len(door_col):
                b = door_col[i]
                c = ws.cell(row=row_offset, column=2, value=b['Bloco'])
                c.fill = fill
                c.border = thick_border
                c.alignment = Alignment(horizontal='center')
                
            if i < len(blind_col):
                b = blind_col[i]
                c = ws.cell(row=row_offset, column=3, value=b['Bloco'])
                c.fill = fill
                c.border = thick_border
                c.alignment = Alignment(horizontal='center')
                
            row_offset += 1
            
        door_len = max([b['l_eff'] for b in door_col]) if door_col else 0
        blind_len = max([b['l_eff'] for b in blind_col]) if blind_col else 0
        
        ws.cell(row=row_offset, column=2, value=round(door_len, 2)).alignment = Alignment(horizontal='center')
        ws.cell(row=row_offset, column=3, value=round(blind_len, 2)).alignment = Alignment(horizontal='center')
        ws.cell(row=row_offset, column=4, value=round(door_len + blind_len, 2)).alignment = Alignment(horizontal='center')
        row_offset += 1
        
        door_w = sum([b['TONS'] for b in door_col])
        blind_w = sum([b['TONS'] for b in blind_col])
        total_door_w = cont.get_total_col_weight(0)
        total_blind_w = cont.get_total_col_weight(1)
        
        ws.cell(row=row_offset, column=2, value=round(total_door_w, 3)).alignment = Alignment(horizontal='center')
        ws.cell(row=row_offset, column=3, value=round(total_blind_w, 3)).alignment = Alignment(horizontal='center')
        ws.cell(row=row_offset, column=4, value=round(total_door_w + total_blind_w, 3)).alignment = Alignment(horizontal='center')
        row_offset += 3
        
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

uploaded_file = st.file_uploader("Selecione sua planilha de blocos (Excel)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name='BLOCOS')
        df_blocks = df.iloc[5:37, [0, 1, 2, 3, 4]].copy()
        df_blocks.columns = ['Bloco', 'Comp', 'Alt', 'Larg', 'TONS']
        for col in ['Comp', 'Alt', 'Larg', 'TONS']:
            df_blocks[col] = pd.to_numeric(df_blocks[col], errors='coerce')
        df_blocks = df_blocks.dropna(subset=['Bloco', 'TONS'])
        blocks_data = df_blocks.to_dict('records')
        
        st.success(f"Planilha carregada! {len(blocks_data)} blocos encontrados.")
        
        if st.button("Gerar Plano de Carregamento", type="primary"):
            with st.spinner("Calculando a melhor distribuição equilibrada..."):
                sol = solve(blocks_data)
                
            if sol:
                st.success("Plano gerado com sucesso!")
                st.subheader("Resumo do Carregamento:")
                for c in sol:
                    st.write(f"**Container {c.id}:** Peso Total {c.get_total_weight():.3f} ton | Porta: {c.get_total_col_weight(0):.3f}t | Cego: {c.get_total_col_weight(1):.3f}t")
                
                excel_data = create_excel_report(sol)
                st.download_button(
                    label="📥 Baixar Excel do Plano (Para o Pátio)",
                    data=excel_data,
                    file_name="Plano_Alocacao_Equilibrado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Não foi possível encontrar uma combinação.")
                
    except Exception as e:
        st.error(f"Erro ao ler a planilha. Detalhe: {e}")
