import streamlit as st
import sqlite3
import pandas as pd
import os
from datetime import datetime
from gerador_relatorio import gerar_docx 

# Configuração da Página
st.set_page_config(page_title="EcoVerde - RIC", layout="wide", page_icon="🌿")

# --- CSS: DESIGN COMPACTO E LIMPO ---
st.markdown("""
    <style>
    /* Remove padding excessivo do topo */
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    
    /* Botões Verdes EcoVerde */
    div.stButton > button { 
        background-color: #548235 !important; 
        color: white !important; 
        border-radius: 6px;
        border: none;
        height: 45px;
        font-weight: bold;
    }
    div.stButton > button:hover { background-color: #406328 !important; }
    
    /* Rodapé */
    .footer { position: fixed; bottom: 10px; right: 15px; color: #888; font-size: 12px; }
    </style>
    <div class="footer">Criado por <b>Bruno Maia</b></div>
""", unsafe_allow_html=True)

# --- BANCO DE DADOS ---
def inicializar_banco():
    conn = sqlite3.connect('dados_ric_oficial.db')
    c = conn.cursor()
    # Tabelas
    c.execute('''CREATE TABLE IF NOT EXISTS empresas (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT, data_criacao TEXT)''')
    
    # Tenta criar a tabela ruas. Se já existir, ignoramos o erro neste momento (o RESET vai resolver)
    try:
        c.execute('''CREATE TABLE IF NOT EXISTS ruas (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT, classificacao_viaria TEXT, sentido_via TEXT, num_faixas INTEGER, largura_via REAL, tem_pavimento TEXT, tipo_pavimento TEXT, estado_pavimento TEXT, problemas_pavimento TEXT, tem_faixa_estacionamento TEXT, estacionamento_irregular TEXT, local_estacionamento TEXT, tem_calcada TEXT, largura_calcada REAL, classificacao_calcada TEXT, estado_calcada TEXT, tem_acessibilidade TEXT, tipos_acessibilidade TEXT, problemas_calcada TEXT, sinalizacao_vertical TEXT, sinalizacao_horizontal TEXT)''')
    except: pass 
    
    c.execute('''CREATE TABLE IF NOT EXISTS empresa_ruas (id INTEGER PRIMARY KEY AUTOINCREMENT, empresa_id INTEGER, rua_id INTEGER)''')
    c.execute('''CREATE TABLE IF NOT EXISTS peds (id INTEGER PRIMARY KEY AUTOINCREMENT, empresa_id INTEGER, numero_ped INTEGER, distancia_empreendimento REAL, classificacao_tabela TEXT, tem_abrigo TEXT, tipo_abrigo TEXT, condicao_abrigo TEXT, tem_assento TEXT, tipo_assento TEXT, condicao_assento TEXT, linhas_onibus TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS ped_ruas (id INTEGER PRIMARY KEY AUTOINCREMENT, ped_id INTEGER, rua_id INTEGER)''')
    conn.commit(); conn.close()

def limpar_banco_completo():
    conn = sqlite3.connect('dados_ric_oficial.db')
    try:
        conn.execute("DELETE FROM ped_ruas")
        conn.execute("DELETE FROM peds")
        conn.execute("DELETE FROM empresa_ruas")
        conn.execute("DELETE FROM ruas")
        conn.execute("DELETE FROM empresas")
        conn.execute("DELETE FROM sqlite_sequence")
        # Força recriação das tabelas dropando elas
        conn.execute("DROP TABLE IF EXISTS ruas") 
        conn.commit()
        return True
    except Exception as e:
        st.error(f"Erro: {e}")
        return False
    finally: conn.close()

inicializar_banco()

# --- FUNÇÕES ---
def get_conexao(): return sqlite3.connect('dados_ric_oficial.db')

# --- FUNÇÃO CRIAR EMPRESA (RESTAURADA) ---
def criar_empresa(nome):
    conn = get_conexao()
    conn.execute("INSERT INTO empresas (nome, data_criacao) VALUES (?, ?)", 
                 (nome, datetime.now().strftime("%Y-%m-%d")))
    conn.commit()
    conn.close()

def importar_rua(empresa_id, rua_id):
    conn = get_conexao()
    check = pd.read_sql(f"SELECT * FROM empresa_ruas WHERE empresa_id={empresa_id} AND rua_id={rua_id}", conn)
    if check.empty:
        conn.execute("INSERT INTO empresa_ruas (empresa_id, rua_id) VALUES (?, ?)", (empresa_id, rua_id))
        conn.commit()
        conn.close()
        return True
    conn.close()
    return False

def classificar_calcada(largura):
    return "Estreita (Inferior a 1,20m)" if largura < 1.2 else "Ampla (Superior a 1,20m)"

def calcular_class_ped(dist):
    if dist < 100: return "Excelente (< 100m)", "green"
    if dist < 200: return "Ótimo (100-200m)", "green"
    if dist < 400: return "Bom (200-400m)", "blue"
    if dist < 600: return "Regular (400-600m)", "orange"
    if dist < 1000: return "Ruim (600-1000m)", "red"
    return "Péssimo (> 1000m)", "red"

# --- CALLBACKS ---
def callback_salvar_rua(empresa_id):
    if not st.session_state.r_nome: st.error("Nome obrigatório!"); return
    
    obs_p_pav = ", ".join(st.session_state.r_probs_pav) if 'r_probs_pav' in st.session_state else ""
    obs_p_calc = ", ".join(st.session_state.r_probs_calc) if 'r_probs_calc' in st.session_state else ""
    obs_acess = ", ".join(st.session_state.r_itens_acess) if 'r_itens_acess' in st.session_state else ""
    
    tipo_pav_final = st.session_state.r_tipo_pav_sim if st.session_state.r_tem_pav == "Sim" else st.session_state.r_tipo_pav_nao
    
    dados = (
        st.session_state.r_nome, st.session_state.r_classe, st.session_state.r_sentido, 
        st.session_state.r_faixas, st.session_state.r_largura, st.session_state.r_tem_pav, 
        tipo_pav_final, st.session_state.r_estado_pav, obs_p_pav,
        st.session_state.r_faixa_est, st.session_state.r_estacionam, st.session_state.r_local_est,
        st.session_state.r_tem_calc, st.session_state.r_larg_calc, classificar_calcada(st.session_state.r_larg_calc),
        st.session_state.r_estado_calc, st.session_state.r_acess, obs_acess, obs_p_calc,
        st.session_state.r_sin_v, st.session_state.r_sin_h
    )
    
    conn = get_conexao(); c = conn.cursor()
    c.execute('''INSERT INTO ruas (nome, classificacao_viaria, sentido_via, num_faixas, largura_via, tem_pavimento, tipo_pavimento, estado_pavimento, problemas_pavimento, tem_faixa_estacionamento, estacionamento_irregular, local_estacionamento, tem_calcada, largura_calcada, classificacao_calcada, estado_calcada, tem_acessibilidade, tipos_acessibilidade, problemas_calcada, sinalizacao_vertical, sinalizacao_horizontal) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', dados)
    rua_id = c.lastrowid
    c.execute("INSERT INTO empresa_ruas (empresa_id, rua_id) VALUES (?, ?)", (empresa_id, rua_id))
    conn.commit(); conn.close()
    
    st.session_state.r_nome = ""
    st.session_state.r_largura = 0.0
    st.toast("Rua Salva com Sucesso!", icon="✅")

def callback_salvar_ped(empresa_id, lista_ruas_dict):
    if not st.session_state.p_ruas: st.error("Selecione ruas!"); return
    
    ids = [lista_ruas_dict[r] for r in st.session_state.p_ruas]
    class_txt, _ = calcular_class_ped(st.session_state.p_dist)
    
    tipo_abrigo = st.session_state.p_tipo_abrigo if st.session_state.p_tem_abrigo == "Sim" else "N/A"
    cond_abrigo = st.session_state.p_cond_abrigo if st.session_state.p_tem_abrigo == "Sim" else "N/A"
    tipo_assento = st.session_state.p_tipo_assento if st.session_state.p_tem_assento == "Sim" else "N/A"
    cond_assento = st.session_state.p_cond_assento if st.session_state.p_tem_assento == "Sim" else "N/A"

    dados = (
        empresa_id, st.session_state.p_num, st.session_state.p_dist, class_txt,
        st.session_state.p_tem_abrigo, tipo_abrigo, cond_abrigo,
        st.session_state.p_tem_assento, tipo_assento, cond_assento,
        st.session_state.p_linhas
    )
    
    conn = get_conexao(); c = conn.cursor()
    c.execute('''INSERT INTO peds (empresa_id, numero_ped, distancia_empreendimento, classificacao_tabela, tem_abrigo, tipo_abrigo, condicao_abrigo, tem_assento, tipo_assento, condicao_assento, linhas_onibus) VALUES (?,?,?,?,?,?,?,?,?,?,?)''', dados)
    pid = c.lastrowid
    for rid in ids: c.execute("INSERT INTO ped_ruas (ped_id, rua_id) VALUES (?,?)", (pid, rid))
    conn.commit(); conn.close()
    
    st.session_state.p_dist = 0.0
    st.session_state.p_num += 1
    st.session_state.p_linhas = ""
    st.toast("PED Salvo!", icon="✅")

# --- UI PRINCIPAL ---
with st.sidebar:
    if os.path.exists("logo.png"): st.image("logo.png", use_column_width=True)
    else: st.markdown("<h2 style='color:#548235;text-align:center'>🌿 EcoVerde</h2>", unsafe_allow_html=True)
    
    df_e = pd.read_sql("SELECT * FROM empresas ORDER BY id DESC", get_conexao())
    empresa_id = None
    empresa_nome = ""
    
    if not df_e.empty:
        opcoes = {f"{r['id']} - {r['nome']}": r['id'] for i, r in df_e.iterrows()}
        sel = st.selectbox("Projeto Atual:", list(opcoes.keys()))
        empresa_id = opcoes[sel]
        empresa_nome = sel.split(" - ")[1]
    
    menu = st.radio("Navegação", ["Nova Empresa", "Gerenciar Ruas", "Cadastro PEDs", "Relatório e Download"])
    
    st.markdown("---")
    st.write(" ")
    if st.button("⚠️ LIMPAR BANCO DE DADOS"):
        if limpar_banco_completo(): 
            inicializar_banco()
            st.success("Banco limpo! Pode cadastrar.")
            st.rerun()

# --- TELA 1: NOVA EMPRESA ---
if menu == "Nova Empresa":
    st.title("🏢 Novo Projeto")
    # CORREÇÃO DE ALINHAMENTO: vertical_alignment="bottom" (Alinha o botão com o input)
    c1, c2 = st.columns([3, 1], vertical_alignment="bottom")
    with c1:
        nome = st.text_input("Nome do Empreendimento / Cliente")
    with c2:
        if st.button("Criar Projeto", use_container_width=True):
            if nome: criar_empresa(nome); st.success("Criado!"); st.rerun()

# --- TELA 2: RUAS ---
elif menu == "Gerenciar Ruas" and empresa_id:
    st.title(f"🛣️ Ruas: {empresa_nome}")
    
    t1, t2 = st.tabs(["➕ Cadastro", "🔍 Banco Global"])
    with t1:
        st.subheader("1. Características e Geometria")
        c1, c2, c3 = st.columns(3)
        c1.text_input("Nome da Via", key="r_nome")
        c2.selectbox("Classificação Viária", ["Local", "Coletora", "Arterial", "Rodovia"], key="r_classe")
        c3.selectbox("Sentido de Circulação", ["Mão Dupla", "Mão Única"], key="r_sentido")
        
        c4, c5, c6 = st.columns(3)
        c4.number_input("Largura da Pista (m)", 0.0, step=0.5, key="r_largura")
        c5.number_input("Nº Faixas de Rolamento", 1, step=1, key="r_faixas")
        c6.selectbox("Possui Faixa de Estacionamento?", ["Não", "Sim (Ambos os lados)", "Sim (Um lado)"], key="r_faixa_est")
        
        st.markdown("---")
        st.subheader("2. Pavimentação e Estacionamento")
        
        col_est1, col_est2 = st.columns(2)
        estacionam = col_est1.selectbox("Usuários estacionam na via?", ["Sim", "Não"], key="r_estacionam")
        if estacionam == "Sim":
            # CORREÇÃO: Opções específicas que você pediu
            col_est2.selectbox("Situação do Estacionamento?", 
                               ["Estacionam em local adequado", 
                                "Estacionam em local não adequado", 
                                "Estacionam em cima da calçada de forma irregular"], 
                               key="r_local_est")
        else:
            st.session_state.r_local_est = "N/A"
        
        col_pav1, col_pav2, col_pav3 = st.columns([1, 2, 2])
        col_pav1.radio("Possui Pavimento?", ["Sim", "Não"], horizontal=True, key="r_tem_pav")
        
        with col_pav2:
            if st.session_state.r_tem_pav == "Sim":
                st.selectbox("Tipo Pavimento", ["Asfalto", "Bloquete", "Concreto", "Poliedrico"], key="r_tipo_pav_sim")
            else:
                st.selectbox("Tipo Revestimento", ["Terra", "Cascalho", "Leito Natural"], key="r_tipo_pav_nao")
        
        with col_pav3:
            st.selectbox("Estado de Conservação", ["Bom", "Regular", "Ruim", "Péssimo"], key="r_estado_pav")
        
        if st.session_state.r_estado_pav != "Bom":
            st.multiselect("Problemas do Pavimento:", ["Buracos", "Trincas/Rachaduras", "Couro de Jacaré", "Remendos", "Afundamentos", "Desgaste excessivo"], key="r_probs_pav")
        
        st.markdown("---")
        st.subheader("3. Calçada e Acessibilidade")
        
        c_calc1, c_calc2, c_calc3 = st.columns(3)
        c_calc1.selectbox("Possui Calçada?", ["Ambos os lados", "Apenas um lado", "Inexistente"], key="r_tem_calc")
        
        if st.session_state.r_tem_calc != "Inexistente":
            c_calc2.number_input("Largura Calçada (m)", 0.0, step=0.1, key="r_larg_calc")
            c_calc3.selectbox("Estado Calçada", ["Bom", "Regular", "Ruim", "Péssimo"], key="r_estado_calc")
            
            if st.session_state.r_estado_calc != "Bom":
                st.multiselect("Problemas na Calçada/Trajeto:", ["Vegetação Alta", "Buracos", "Desníveis/Degraus", "Obstáculos (Postes)", "Piso Escorregadio", "Trincas", "Estreitamento"], key="r_probs_calc")
            
            c_ac1, c_ac2 = st.columns([1, 3])
            c_ac1.radio("Possui Acessibilidade?", ["Sim", "Não"], horizontal=True, key="r_acess")
            if st.session_state.r_acess == "Sim":
                st.multiselect("Itens de Acessibilidade:", ["Rampa de Acesso", "Piso Tátil Alerta", "Piso Tátil Direcional", "Rebaixamento de Guia"], key="r_itens_acess")
        
        st.markdown("---")
        st.subheader("4. Sinalização")
        cs1, cs2 = st.columns(2)
        cs1.selectbox("Sinalização Vertical", ["Existente/Boa", "Existente/Ruim", "Inexistente"], key="r_sin_v")
        cs2.selectbox("Sinalização Horizontal", ["Existente/Boa", "Existente/Ruim", "Inexistente"], key="r_sin_h")

        st.write("")
        st.button("💾 SALVAR RUA NO PROJETO", on_click=callback_salvar_rua, args=(empresa_id,), type="primary")

    with t2:
        st.subheader("Importar ruas de outros projetos")
        st.info("Aqui aparecem ruas cadastradas em outros projetos para você reutilizar.")
        
        conn = get_conexao()
        q = f"SELECT * FROM ruas WHERE id NOT IN (SELECT rua_id FROM empresa_ruas WHERE empresa_id={empresa_id})"
        df_glob = pd.read_sql(q, conn)
        conn.close()
        
        if not df_glob.empty:
            c_imp1, c_imp2 = st.columns([3, 1], vertical_alignment="bottom")
            with c_imp1:
                sel = st.selectbox("Selecione a rua para importar:", df_glob['nome'] + " (ID: " + df_glob['id'].astype(str) + ")")
            with c_imp2:
                if st.button("📥 Importar Rua"):
                    id_imp = int(sel.split("ID: ")[1].replace(")",""))
                    if importar_rua(empresa_id, id_imp):
                        st.success("Rua importada com sucesso!")
                        st.rerun()
        else:
            st.warning("Não há ruas disponíveis no banco global para importar.")

# --- TELA 3: PEDS ---
elif menu == "Cadastro PEDs" and empresa_id:
    st.title(f"🚏 PEDs: {empresa_nome}")
    
    c_top1, c_top2 = st.columns(2)
    with c_top1:
        st.number_input("Número do PED", min_value=1, key="p_num")
        st.number_input("Distância do Empreendimento (m)", 0.0, step=10.0, key="p_dist")
        
        txt, color = calcular_class_ped(st.session_state.p_dist)
        if color == "green": st.success(f"Classificação: {txt}")
        elif color == "red": st.error(f"Classificação: {txt}")
        else: st.info(f"Classificação: {txt}")

    with c_top2:
        st.text_input("Linhas de Ônibus", key="p_linhas")
        
        conn = get_conexao()
        ruas = pd.read_sql(f"SELECT r.id, r.nome FROM ruas r JOIN empresa_ruas er ON r.id=er.rua_id WHERE er.empresa_id={empresa_id}", conn)
        conn.close()
        
        ruas_dict = dict(zip(ruas['nome'], ruas['id']))
        st.multiselect("Ruas do Trajeto", list(ruas_dict.keys()), key="p_ruas")

    st.markdown("---")
    st.subheader("Infraestrutura do Ponto")
    
    c_infra1, c_infra2 = st.columns(2)
    with c_infra1:
        st.radio("Possui Abrigo?", ["Sim", "Não"], horizontal=True, key="p_tem_abrigo")
        if st.session_state.p_tem_abrigo == "Sim":
            st.selectbox("Tipo de Abrigo", ["Concreto", "Metálico", "Vidro", "Madeira"], key="p_tipo_abrigo")
            st.selectbox("Estado do Abrigo", ["Bom", "Regular", "Ruim", "Danificado"], key="p_cond_abrigo")
    
    with c_infra2:
        st.radio("Possui Assento?", ["Sim", "Não"], horizontal=True, key="p_tem_assento")
        if st.session_state.p_tem_assento == "Sim":
            st.selectbox("Tipo de Assento", ["Concreto", "Metálico", "Madeira", "Plástico"], key="p_tipo_assento")
            st.selectbox("Estado do Assento", ["Bom", "Regular", "Ruim", "Danificado"], key="p_cond_assento")
            
    st.write("")
    st.button("💾 SALVAR PED", on_click=callback_salvar_ped, args=(empresa_id, ruas_dict), type="primary")

# --- TELA 4: RELATÓRIO E DADOS ---
elif menu == "Relatório e Download" and empresa_id:
    st.title("📊 Download")
    if st.button("📄 GERAR RELATÓRIO WORD", use_container_width=True):
        arq = gerar_docx(empresa_id, empresa_nome)
        with open(arq, "rb") as f:
            st.download_button("⬇️ BAIXAR ARQUIVO", f, file_name=arq, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    
    st.divider()
    st.subheader("📋 Ruas Cadastradas")
    df_ruas = pd.read_sql(f"SELECT nome, classificacao_viaria, tipo_pavimento, estado_pavimento FROM ruas r JOIN empresa_ruas er ON r.id=er.rua_id WHERE er.empresa_id={empresa_id}", get_conexao())
    st.dataframe(df_ruas, use_container_width=True)
    
    st.subheader("📋 PEDs Cadastrados")
    df_peds = pd.read_sql(f"SELECT numero_ped, distancia_empreendimento, classificacao_tabela FROM peds WHERE empresa_id={empresa_id}", get_conexao())
    st.dataframe(df_peds, use_container_width=True)