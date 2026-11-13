import streamlit as st
import pandas as pd
import openpyxl
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import json
import os

# --- CONFIGURAÇÕES DA PÁGINA ---
st.set_page_config(
    page_title="Gerador de Mapa de Entrega - Sua Música",
    page_icon="📊",
    layout="centered"
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Mapeamento De -> Para (Arquivo 1 -> Aba Aux do Template)
DE_PARA_FORMATOS = {
    "Combo (1 Post + 1 Sequência de stories (Perfil Influenciador M)": "Combo (1 Post + 1 Sequências de stories (Perfil Influenciador M)",
    "Combo (1 Post + 1 Sequência de stories (Perfil Influenciador P)": "Combo (1 Post + 1 Sequências de stories (Perfil Influenciador P)",
    'Video Pre-roll 30" CPV': "Video Pre-roll CPV",
    "Ação de merchandising": "Ação de Merchandising",
}

def normalizar_formato(nome: str) -> str:
    nome_limpo = str(nome).strip()
    return DE_PARA_FORMATOS.get(nome_limpo, nome_limpo)

def ler_arquivo1(file_buffer):
    wb = openpyxl.load_workbook(file_buffer, data_only=True)
    ws = wb["Mídia"] if "Mídia" in wb.sheetnames else wb.active

    # Extração dinâmica dos metadados do cabeçalho do Excel
    campanha = ws["C6"].value or ""
    cliente = ws["C3"].value or ""
    agencia = ws["C5"].value or ""
    executivo = ws["C4"].value or ""
    mes = ws["G6"].value or ""
    data_veiculacao = ws["G5"].value or ""

    itens = []
    header_encontrado = False
    
    # Índices padrão (caso não ache no cabeçalho)
    col_formato = 2  # Coluna B
    col_unidade = 4  # Coluna D
    col_qntd = 6     # Coluna F
    col_tabela = 7   # Coluna G
    col_desconto = 9 # Padrão anterior (I)

    for r in range(1, ws.max_row + 1):
        formato_val = ws.cell(r, 2).value
        
        # Procura dinamicamente a linha de cabeçalho da tabela de mídia
        if str(formato_val).strip() == "Formato":
            header_encontrado = True
            # Varre os cabeçalhos desta linha para achar a coluna exata de "Desconto Aplicado" e "Qtd"
            for col_idx in range(1, ws.max_column + 1):
                txt_cabecalho = str(ws.cell(r, col_idx).value or "").strip().lower()
                if "desconto" in txt_cabecalho or "desc" in txt_cabecalho:
                    col_desconto = col_idx
                elif txt_cabecalho in ["qntd", "quantidade", "qtd"]:
                    col_qntd = col_idx
            continue

        if not header_encontrado or formato_val is None:
            continue

        formato_str = str(formato_val).strip()
        if formato_str.lower() in ["total", "observações gerais"]:
            break

        unidade = ws.cell(r, col_unidade).value
        qntd = ws.cell(r, col_qntd).value
        tabela = ws.cell(r, col_tabela).value
        desconto = ws.cell(r, col_desconto).value

        if qntd is None and tabela is None:
            itens.append({
                "tipo": "secao",
                "formato": formato_str
            })
        else:
            itens.append({
                "tipo": "item",
                "formato": normalizar_formato(formato_str),
                "unidade": str(unidade).strip() if unidade else "",
                "qntd": float(qntd) if qntd is not None else 0,
                "desconto": float(desconto) if desconto is not None else 0.0
            })

    meta_extraida = {
        "cliente": str(cliente),
        "agencia": str(agencia),
        "executivo": str(executivo),
        "mes": str(mes),
        "campanha": str(campanha),
        "data_veiculacao": str(data_veiculacao)
    }

    return meta_extraida, itens

def conectar_google():
    caminho_json = "credentials.json"
    if not os.path.exists(caminho_json):
        st.error("Arquivo 'credentials.json' não encontrado na pasta do projeto!")
        st.stop()
    creds = Credentials.from_service_account_file(caminho_json, scopes=SCOPES)
        
    gc = gspread.authorize(creds)
    drive_service = build("drive", "v3", credentials=creds)
    return gc, drive_service

def gerar_planilha_sheets(gc, drive_service, template_id, pasta_id, meta, itens):
    # Nome atualizado para conter "Mapa", o Cliente, a Campanha e o Mês
    nome_arquivo = f"Mapa - {meta['cliente']} - {meta['campanha']} - {meta['mes']}"
    
    # 1. Faz a cópia direta do template no Drive Compartilhado
    body = {
        "name": nome_arquivo,
    }
    if pasta_id:
        body["parents"] = [pasta_id]

    novo_arquivo = drive_service.files().copy(
        fileId=template_id,
        body=body,
        supportsAllDrives=True
    ).execute()
    
    nova_planilha_id = novo_arquivo["id"]
    
    # 2. Abre a planilha recém-copiada
    sh = gc.open_by_key(nova_planilha_id)
    ws_mapa = sh.worksheet("Mapa Entrega Total")
    ws_resumo = sh.worksheet("Resumo Entrega Total")

    # 3. Busca dinâmica da célula contendo "Entrega total" nas primeiras 15 linhas
    celula_entrega = None
    todas_linhas = ws_mapa.get("A1:AZ15")
    
    linha_cabecalho = 9
    col_entrega_idx = 34
    
    for r_idx, linha in enumerate(todas_linhas, start=1):
        for c_idx, valor in enumerate(linha, start=1):
            if str(valor).strip().lower() == "entrega total":
                celula_entrega = (r_idx, c_idx)
                break
        if celula_entrega:
            break
            
    if celula_entrega:
        linha_cabecalho = celula_entrega[0]
        col_entrega_idx = celula_entrega[1]
        
    def numero_para_letra(num):
        letra = ""
        while num > 0:
            num, remainder = divmod(num - 1, 26)
            letra = chr(65 + remainder) + letra
        return letra

    col_entrega_letra = numero_para_letra(col_entrega_idx)
    
    col_desc_letra = None
    if linha_cabecalho <= len(todas_linhas):
        for c_idx, valor in enumerate(todas_linhas[linha_cabecalho - 1], start=1):
            txt = str(valor).strip().lower()
            if "desc" in txt or "desconto" in txt:
                col_desc_letra = numero_para_letra(c_idx)
                break
                
    if not col_desc_letra:
        col_desc_letra = numero_para_letra(col_entrega_idx + 2)
            
    # Os dados do plano de mídia começam na linha 11 (deixando a linha 10 do template totalmente intacta com suas fórmulas)
    linha_inicial = 11

    # 4. Preenche Cabeçalho do Mapa
    ws_mapa.update("H3", [[meta["cliente"]]])
    ws_mapa.update("H4", [[meta["executivo"]]])
    ws_mapa.update("H5", [[meta["agencia"]]])
    ws_mapa.update("R3", [[meta["campanha"]]])
    ws_mapa.update("R5", [[meta["data_veiculacao"]]])

    # 5. Localiza exatamente onde está a linha "TOTAL" na coluna B ou C para preservar o rodapé e gerenciar o miolo
    coluna_b_valores = ws_mapa.col_values(2)
    coluna_c_valores = ws_mapa.col_values(3)
    linha_total_template = None
    max_linhas_busca = max(len(coluna_b_valores), len(coluna_c_valores))

    for idx in range(linha_inicial, max_linhas_busca + 1):
        val_b = str(coluna_b_valores[idx - 1] if idx <= len(coluna_b_valores) else "").strip().lower()
        val_c = str(coluna_c_valores[idx - 1] if idx <= len(coluna_c_valores) else "").strip().lower()
        if "total" in val_b or "total" in val_c:
            linha_total_template = idx
            break
            
    if not linha_total_template:
        linha_total_template = 26  # Linha padrão do total no template

    qtd_itens_novos = len(itens)
    linha_fim_dados_template = linha_total_template - 1
    qtd_linhas_template = linha_fim_dados_template - linha_inicial + 1
    sheet_id = ws_mapa.id

    # 6. Ajusta o miolo da tabela (remove apenas as linhas vazias excedentes antes do total, ou insere se faltar)
    if qtd_itens_novos != qtd_linhas_template:
        diff = qtd_itens_novos - qtd_linhas_template
        if diff > 0:
            body_insert = {
                "requests": [{
                    "insertDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": linha_fim_dados_template - 1,
                            "endIndex": linha_fim_dados_template - 1 + diff
                        },
                        "inheritFromBefore": True
                    }
                }]
            }
            sh.batch_update(body_insert)
        elif diff < 0:
            body_delete = {
                "requests": [{
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": linha_inicial + qtd_itens_novos - 1,
                            "endIndex": linha_total_template - 1
                        }
                    }
                }]
            }
            sh.batch_update(body_delete)

    # 7. Preenche os itens e seções a partir da linha 11, sem tocar na linha 10 e preservando as fórmulas nativas
    linha_atual = linha_inicial
    batch_updates = []
    linhas_secao = []

    for entry in itens:
        r = linha_atual
        if entry["tipo"] == "secao":
            batch_updates.append({"range": f"B{r}", "values": [[entry["formato"]]]})
            ws_mapa.batch_clear([f"C{r}:AP{r}"])
            linhas_secao.append(r)
        else:
            batch_updates.extend([
                {"range": f"B{r}", "values": [[entry["formato"]]]},
                {"range": f"{col_entrega_letra}{r}", "values": [[entry["qntd"]]]},
                {"range": f"{col_desc_letra}{r}", "values": [[entry["desconto"]]]}
            ])
        linha_atual += 1

    ws_mapa.batch_update(batch_updates, value_input_option="USER_ENTERED")

    # 8. Formatação das linhas de seção: remove dropdown e aplica o estilo visual (fundo azul escuro e texto branco)
    if linhas_secao:
        sheet_id = ws_mapa.id
        requests_secao = []
        for r in linhas_secao:
            requests_secao.append({
                "setDataValidation": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": r - 1,
                        "endRowIndex": r,
                        "startColumnIndex": 1,
                        "endColumnIndex": 2
                    }
                }
            })
            requests_secao.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": r - 1,
                        "endRowIndex": r,
                        "startColumnIndex": 1,
                        "endColumnIndex": 42
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.08, "green": 0.21, "blue": 0.36},
                            "textFormat": {
                                "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                                "bold": True
                            },
                            "borders": {
                                "top": {"style": "NONE"},
                                "bottom": {"style": "NONE"},
                                "left": {"style": "NONE"},
                                "right": {"style": "NONE"}
                            }
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,borders)"
                }
            })
        sh.batch_update({"requests": requests_secao})

    return f"https://docs.google.com/spreadsheets/d/{nova_planilha_id}/edit"

# --- INTERFACE DA APLICAÇÃO ---
st.title("📊 Gerador de Mapa de Entrega")
st.markdown("Faça o upload do **Plano de Mídia (.xlsx)** para gerar o Google Sheets preenchido automaticamente.")

arquivo1 = st.file_uploader("Enviar Plano de Mídia", type=["xlsx"])

if arquivo1:
    meta_extraida, itens = ler_arquivo1(arquivo1)
    st.success(f"Arquivo lido com sucesso! {sum(1 for i in itens if i['tipo']=='item')} formatos válidos encontrados.")

    with st.form("form_mapa"):
        col1, col2 = st.columns(2)
        cliente = col1.text_input("Cliente", value=meta_extraida["cliente"])
        agencia = col2.text_input("Agência", value=meta_extraida["agencia"])
        executivo = col1.text_input("Executivo", value=meta_extraida["executivo"])
        mes = col2.text_input("Mês de Veiculação", value=meta_extraida["mes"])
        campanha = st.text_input("Campanha / PI", value=meta_extraida["campanha"])
        data_veiculacao = st.text_input("Período de Veiculação", value=meta_extraida["data_veiculacao"])

        gerar_btn = st.form_submit_button("🚀 Gerar Planilha no Google Sheets")

    if gerar_btn:
        with st.spinner("Criando cópia do template e preenchendo os dados..."):
            try:
                gc, drive_service = conectar_google()
                
                template_id = "1p_I2qvTjrntUCp7992Z7W1TPbUJmCXHGWwNALalSnHs"
                pasta_id = ""

                meta = {
                    "cliente": cliente,
                    "agencia": agencia,
                    "executivo": executivo,
                    "mes": mes,
                    "campanha": campanha,
                    "data_veiculacao": data_veiculacao,
                }
                
                url_planilha = gerar_planilha_sheets(gc, drive_service, template_id, pasta_id, meta, itens)
                
                st.balloons()
                st.success("Mapa de Entrega gerado com sucesso!")
                st.markdown(f"### 👉 **[Clique aqui para abrir a planilha no Google Sheets]({url_planilha})**")
            except Exception as e:
                st.error(f"Ocorreu um erro ao gerar a planilha: {e}")