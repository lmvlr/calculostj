import os
import sys
import pandas as pd
from datetime import datetime
from pandas.tseries.offsets import DateOffset
from flask import Flask, request, send_file, jsonify
from io import BytesIO

# Bibliotecas para PDF
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import cm

# Habilitar CORS
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # Libera as requisições de outros domínios

# #######################################
# Caminhos das tabelas
# #######################################
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
caminho_log = os.path.join(BASE_DIR, "log.txt")

arquivo_tabela_pratica = os.path.join(BASE_DIR, "Tabela_Pratica_TJSP.xlsx")
arquivo_tabela_ipcae   = os.path.join(BASE_DIR, "Tabela_IPCA-E.xlsx")
arquivo_tabela_selic   = os.path.join(BASE_DIR, "Tabela_Selic.xlsx")
arquivo_selic_antes    = os.path.join(BASE_DIR, "Selic_antes_2022.xlsx")

# #######################################
# Ler as planilhas
# #######################################
tabela_selic_antes = pd.read_excel(arquivo_selic_antes)
tabela_selic_antes["PERÍODO DE VIGÊNCIA INICIAL"] = pd.to_datetime(tabela_selic_antes["PERÍODO DE VIGÊNCIA INICIAL"])
tabela_selic_antes["PERÍODO DE VIGÊNCIA FINAL"]   = pd.to_datetime(tabela_selic_antes["PERÍODO DE VIGÊNCIA FINAL"])

tabela_pratica = pd.read_excel(arquivo_tabela_pratica)
tabela_ipcae   = pd.read_excel(arquivo_tabela_ipcae)
tabela_selic   = pd.read_excel(arquivo_tabela_selic)

for tbl in [tabela_pratica, tabela_ipcae, tabela_selic]:
    tbl["Ano"] = tbl["Ano"].astype(int)
    tbl["Mês"] = tbl["Mês"].astype(int)

tabela_selic.sort_values(by=["Ano", "Mês"], inplace=True, ignore_index=True)

# #######################################
# Funções auxiliares
# #######################################
def br_format(value):
    """Converte float -> str em formato brasileiro (1.234,56)."""
    if pd.isna(value):
        return "-"
    s = f"{value:,.2f}"
    s = s.replace(",", "X")  # "123X456.78"
    s = s.replace(".", ",")  # "123X456,78"
    s = s.replace("X", ".")  # "123.456,78"
    return s

def regra_oc_data(ts: pd.Timestamp):
    """Calcula a OC via data do ofício/cessão (antigamente 'SEM OC')."""
    if pd.isna(ts):
        return None
    y, m = ts.year, ts.month
    if m < 3:
        return y + 1
    else:
        return y + 2

def determina_oc(ativo_row):
    """Tenta pegar 'Ordem Cronológica' ou deduzir de 'Data oficio'."""
    oc_str = ativo_row.get("Ordem Cronológica", "")
    if pd.notna(oc_str) and oc_str.strip().upper() != "SEM OC":
        try:
            splitted = oc_str.split("/")[-1]
            oc_ok = int(float(splitted))
            return (oc_ok, None)
        except:
            return (None, f"Ordem Cronológica inválida: {oc_str}")
    else:
        return (None, "SEM OC + Data ofício ausente (não usado mais)")

# #######################################
# Gera PDF
# #######################################
def gerar_pdf_para_ativo(
    nome_ativo,
    data_base_str,
    final_date,
    historico_normal,
    historico_punitivo,
    valor_total_final,
    ordem_cronologica,
    valores_iniciais_str
):
    """Cria PDF em memória, sem datas de cessão."""
    styles = getSampleStyleSheet()
    story = []

    # Título
    titulo = Paragraph(f"<b>Relatório de Cálculo - {nome_ativo}</b>", styles["Title"])
    story.append(titulo)
    story.append(Spacer(1, 0.3 * cm))

    # Cabeçalho
    head_html = f"""
    <b>VALORES INICIAIS</b><br/>{valores_iniciais_str}<br/><br/>
    <b>Ordem Cronológica:</b> {ordem_cronologica}<br/>
    <b>Data Base:</b> {data_base_str}<br/>
    <b>Período calculado:</b> até {final_date.strftime('%m/%Y')}
    """
    story.append(Paragraph(head_html, styles["Normal"]))
    story.append(Spacer(1, 0.3 * cm))

    # Tabela Resumo (apenas Valor Total)
    mes_ano_final = final_date.strftime("%m/%Y")
    titulo_resumo = f"RESUMO DO ATIVO ATÉ O MÊS {mes_ano_final}".upper()

    table_resumo_data = [
        [titulo_resumo, ""],
        ["Valor Total do Ativo (mês atual)", f"R$ {br_format(valor_total_final)}"],
    ]

    table_resumo = Table(table_resumo_data, colWidths=[8 * cm, 9 * cm])
    style_resumo = TableStyle([
        ("SPAN", (0, 0), (1, 0)),
        ("BACKGROUND", (0, 0), (1, 0), colors.black),
        ("TEXTCOLOR", (0, 0), (1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.black),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.grey)
    ])
    table_resumo.setStyle(style_resumo)
    story.append(table_resumo)
    story.append(Spacer(1, 0.4 * cm))

    # Tabela do histórico
    pun_map = {x["data"]: x["acumulado"] for x in historico_punitivo}
    header = ["Mês/Ano", "Principal", "Desc. Prev", "Desc. Assist", "Juros Normal", "Juros Punitivo", "Soma Juros"]
    tdata = [header]

    datas_hist = sorted(x["data"] for x in historico_normal)
    for dt_ in datas_hist:
        row_ = next(x for x in historico_normal if x["data"] == dt_)
        p_ = row_["Principal Líquido"]
        dp_ = row_["Desconto Previdenciário"]
        da_ = row_["Desconto Assistência médica"]
        jn_ = row_["Juros"]
        pun_ = pun_map.get(dt_, 0.0)
        soma_j = jn_ + pun_
        mes_ano = dt_.strftime("%m/%Y")
        tdata.append([
            mes_ano,
            br_format(p_),
            br_format(dp_),
            br_format(da_),
            br_format(jn_),
            br_format(pun_),
            br_format(soma_j)
        ])

    st = TableStyle([
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
    ])
    colws = [2.2 * cm] * 7
    hist_table = Table(tdata, colWidths=colws)
    hist_table.setStyle(st)
    story.append(hist_table)

    # Rodapé
    def rodape(canv, doc):
        pag = doc.page
        data_hoje = datetime.now().strftime("%d/%m/%Y %H:%M")
        canv.setFont("Helvetica", 8)
        left_txt = "LM Cálculos, com base na atualização DEPRE TJ-SP"
        canv.drawString(2 * cm, 1.1 * cm, left_txt)
        right_txt = f"Página {pag} - {data_hoje}"
        canv.drawRightString(19.5 * cm, 1.1 * cm, right_txt)

    pdf_buffer = BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4)
    doc.build(story, onFirstPage=rodape, onLaterPages=rodape)
    pdf_buffer.seek(0)

    return pdf_buffer

# #######################################
# Rota principal
# #######################################
@app.route("/calcular", methods=["POST"])
def calcular():
    """
    Espera JSON de UM ativo com os campos:
    {
      "Nome Completo": "Fulano",
      "Ordem Cronológica": "2020",
      "Data Base": "2006-06-01",
      "Principal Líquido": 10000.0,
      "Juros": 500.0,
      "Desconto Previdenciário": 200.0,
      "Desconto Assistência médica": 150.0
    }

    Retorna um arquivo PDF com extensão .pdf.
    """
    data = request.get_json()
    if not data or not isinstance(data, dict):
        return jsonify({"error": "Formato JSON inválido. Esperamos um objeto com os campos do ativo"}), 400

    nome_ativo = data.get("Nome Completo", "NOME_NAO_INFORMADO")
    print(f"\n================== Cálculo para {nome_ativo} ==================")

    oc, erro_oc = determina_oc(data)
    if oc is None:
        msg = f"Ordem Cronológica inválida: {erro_oc}"
        print(msg)
        return jsonify({"error": msg}), 400
    else:
        ordem_cronologica = oc

    data_base = pd.to_datetime(data.get("Data Base"), errors="coerce")
    if pd.isna(data_base):
        msg = "Data Base ausente ou inválida"
        print(msg)
        return jsonify({"error": msg}), 400

    val_princ = data.get("Principal Líquido", 0.0) or 0.0
    val_juros = data.get("Juros", 0.0) or 0.0
    val_dp    = data.get("Desconto Previdenciário", 0.0) or 0.0
    val_da    = data.get("Desconto Assistência médica", 0.0) or 0.0

    valores_iniciais_str = (
       f"Principal Líquido: R$ {br_format(val_princ)}<br/>"
       f"Juros: R$ {br_format(val_juros)}<br/>"
       f"Desconto Prev: R$ {br_format(val_dp)}<br/>"
       f"Desconto Assist: R$ {br_format(val_da)}"
    )

    today = pd.Timestamp.today()
    final_date = pd.Timestamp(year=today.year, month=today.month, day=1)
    data_base_start = data_base.replace(day=1)
    print(f"   Data Base: {data_base_start.strftime('%d/%m/%Y')}")
    print(f"   OC: {ordem_cronologica}")

    # Inicializa variáveis de cálculo
    var_names = ["Principal Líquido", "Juros", "Desconto Previdenciário", "Desconto Assistência médica"]
    normal_values = {}
    normal_values[data_base_start] = {v: data.get(v, 0.0) for v in var_names}
    historico_normal_soma = {}
    historico_normal = []

    current_date = data_base_start
    selic_base_norm = {}
    selic_acum_norm = {}

    # 1) Atualização normal
    if ordem_cronologica < 2022:
        inicio_graça = pd.Timestamp(year=ordem_cronologica-1, month=7, day=1)
        fim_graça    = pd.Timestamp(year=ordem_cronologica, month=12, day=31)
    else:
        inicio_graça = pd.Timestamp(year=ordem_cronologica-1, month=4, day=3)
        fim_graça    = pd.Timestamp(year=ordem_cronologica, month=12, day=31)

    while current_date <= final_date:
        if current_date == data_base_start:
            prev_date = current_date
        else:
            prev_date = current_date - DateOffset(months=1)
            prev_date = prev_date.replace(day=1)

        if current_date < pd.Timestamp(2022, 1, 1):
            if inicio_graça <= current_date <= fim_graça:
                if current_date == inicio_graça:
                    new_vals = {}
                    for v_ in var_names:
                        new_vals[v_] = normal_values[prev_date][v_]
                else:
                    fp_arr = tabela_ipcae[
                        (tabela_ipcae["Ano"] == prev_date.year) &
                        (tabela_ipcae["Mês"] == prev_date.month)
                    ]["Índice"].values
                    fc_arr = tabela_ipcae[
                        (tabela_ipcae["Ano"] == current_date.year) &
                        (tabela_ipcae["Mês"] == current_date.month)
                    ]["Índice"].values
                    factor = 1.0
                    if fp_arr.size and fc_arr.size:
                        factor = fc_arr[0] / fp_arr[0]
                    new_vals = {}
                    for v_ in var_names:
                        old_val = normal_values[prev_date][v_]
                        new_vals[v_] = old_val * factor if pd.notna(old_val) else old_val
            else:
                fp_arr = tabela_pratica[
                    (tabela_pratica["Ano"] == prev_date.year) &
                    (tabela_pratica["Mês"] == prev_date.month)
                ]["Índice"].values
                fc_arr = tabela_pratica[
                    (tabela_pratica["Ano"] == current_date.year) &
                    (tabela_pratica["Mês"] == current_date.month)
                ]["Índice"].values
                if fp_arr.size and fc_arr.size:
                    full_factor = fc_arr[0] / fp_arr[0]
                else:
                    full_factor = 1.0
                new_vals = {}
                for v_ in var_names:
                    old_val = normal_values[prev_date][v_]
                    if pd.isna(old_val):
                        new_val = old_val
                    else:
                        if (current_date.year == 2021 and current_date.month == 12 and v_ == "Juros"):
                            partial_factor = 1.0 + (full_factor - 1.0) * (8/31.0)
                            new_val = old_val * partial_factor
                        else:
                            new_val = old_val * full_factor
                    new_vals[v_] = new_val
        else:
            # >= 2022 => SELIC
            new_vals = {}
            for v_ in var_names:
                if v_ not in selic_base_norm:
                    if current_date == data_base_start:
                        selic_base_norm[v_] = normal_values[current_date][v_]
                    else:
                        selic_base_norm[v_] = normal_values[prev_date][v_]
                    selic_acum_norm[v_] = 0.0
                selic_rate_arr = tabela_selic[
                    (tabela_selic["Ano"] == current_date.year) &
                    (tabela_selic["Mês"] == current_date.month)
                ]["Índice"].values
                rate = (selic_rate_arr[0] / 100) if selic_rate_arr.size else 0.0
                juros_mes = selic_base_norm[v_] * rate
                selic_acum_norm[v_] += juros_mes
                new_val = selic_base_norm[v_] + selic_acum_norm[v_]
                new_vals[v_] = new_val

        normal_values[current_date] = new_vals
        p_ = new_vals.get("Principal Líquido", 0.0) or 0.0
        j_ = new_vals.get("Juros", 0.0) or 0.0
        dp_ = new_vals.get("Desconto Previdenciário", 0.0) or 0.0
        da_ = new_vals.get("Desconto Assistência médica", 0.0) or 0.0
        total_ = p_ + j_ + dp_ + da_
        historico_normal_soma[current_date] = total_

        historico_normal.append({
            "data": current_date,
            "Principal Líquido": p_,
            "Juros": j_,
            "Desconto Previdenciário": dp_,
            "Desconto Assistência médica": da_
        })

        if current_date.month == 12:
            current_date = pd.Timestamp(year=current_date.year + 1, month=1, day=1)
        else:
            current_date = pd.Timestamp(year=current_date.year, month=current_date.month + 1, day=1)

    # 2) Juros punitivos
    punitivo_mes = {}
    punitive_accum = 0.0
    if data_base_start > pd.Timestamp(2021, 12, 1):
        print("   Juros punitivos não se aplicam (Data Base > dez/2021).")
    else:
        pun_current_date = data_base_start
        outside_grace_months = 0
        last_punitivo = None
        punitive_base_fixed = None
        fixed_meta_end = pd.Timestamp(2012, 5, 1)

        while pun_current_date <= final_date:
            if pun_current_date == data_base_start:
                pun_prev_date = pun_current_date
            else:
                pun_prev_date = pun_current_date - DateOffset(months=1)
                pun_prev_date = pun_prev_date.replace(day=1)

            if inicio_graça <= pun_current_date <= fim_graça:
                outside_grace_months = 0
                fp_arr = tabela_ipcae[
                    (tabela_ipcae["Ano"] == pun_prev_date.year) &
                    (tabela_ipcae["Mês"] == pun_prev_date.month)
                ]["Índice"].values
                fc_arr = tabela_ipcae[
                    (tabela_ipcae["Ano"] == pun_current_date.year) &
                    (tabela_ipcae["Mês"] == pun_current_date.month)
                ]["Índice"].values
                if pun_current_date == inicio_graça:
                    update_factor = 1.0
                else:
                    if fp_arr.size and fc_arr.size:
                        update_factor = fc_arr[0] / fp_arr[0]
                    else:
                        update_factor = 1.0
                punitive_accum *= update_factor
            else:
                outside_grace_months += 1
                if pun_current_date < pd.Timestamp(2022, 1, 1):
                    vals = normal_values.get(pun_current_date, {})
                    base_punitiva = 0.0
                    if vals:
                        base_punitiva += vals.get("Principal Líquido", 0.0) or 0.0
                        base_punitiva += vals.get("Desconto Previdenciário", 0.0) or 0.0
                        base_punitiva += vals.get("Desconto Assistência médica", 0.0) or 0.0

                    if pun_current_date <= fixed_meta_end:
                        if (pun_current_date.month == fixed_meta_end.month and pun_current_date.year == fixed_meta_end.year):
                            total_rate = 0.005 * outside_grace_months
                            outside_grace_months = 0
                        else:
                            total_rate = 0.0
                    else:
                        selic_rows = tabela_selic_antes[
                            (tabela_selic_antes["PERÍODO DE VIGÊNCIA INICIAL"] <= pun_current_date) &
                            (tabela_selic_antes["PERÍODO DE VIGÊNCIA FINAL"] >= pun_current_date)
                        ]
                        if not selic_rows.empty:
                            meta_end = selic_rows.iloc[0]["PERÍODO DE VIGÊNCIA FINAL"]
                            meta_anual = selic_rows.iloc[0]["META SELIC (A.A) %"]
                            if (pun_current_date.month == meta_end.month and pun_current_date.year == meta_end.year):
                                monthly_rate = 0.70 * ((meta_anual/12)/100) if meta_anual <= 8.5 else 0.005
                                total_rate = monthly_rate * outside_grace_months
                                outside_grace_months = 0
                            else:
                                total_rate = 0.0
                        else:
                            total_rate = 0.005

                    new_increment = base_punitiva * total_rate
                    fp_arr = tabela_pratica[
                        (tabela_pratica["Ano"] == pun_prev_date.year) &
                        (tabela_pratica["Mês"] == pun_prev_date.month)
                    ]["Índice"].values
                    fc_arr = tabela_pratica[
                        (tabela_pratica["Ano"] == pun_current_date.year) &
                        (tabela_pratica["Mês"] == pun_current_date.month)
                    ]["Índice"].values
                    if fp_arr.size and fc_arr.size:
                        update_factor = fc_arr[0] / fp_arr[0]
                    else:
                        update_factor = 1.0
                    punitive_accum = punitive_accum * update_factor + new_increment
                else:
                    if punitive_base_fixed is None:
                        punitive_base_fixed = last_punitivo if last_punitivo else punitive_accum
                    selic_rate_arr = tabela_selic[
                        (tabela_selic["Ano"] == pun_current_date.year) &
                        (tabela_selic["Mês"] == pun_current_date.month)
                    ]["Índice"].values
                    rate = selic_rate_arr[0] / 100 if selic_rate_arr.size else 0.0
                    new_increment = punitive_base_fixed * rate
                    punitive_accum += new_increment

            last_punitivo = punitive_accum
            punitivo_mes[pun_current_date] = punitive_accum

            if pun_current_date.month == 12:
                pun_current_date = pd.Timestamp(year=pun_current_date.year+1, month=1, day=1)
            else:
                pun_current_date = pd.Timestamp(year=pun_current_date.year, month=pun_current_date.month+1, day=1)

    normal_soma_final = historico_normal_soma.get(final_date, 0.0)
    punit_final = punitivo_mes.get(final_date, 0.0)
    valor_total_ativo_final = normal_soma_final + punit_final

    # Gera PDF
    pdf_bytes = gerar_pdf_para_ativo(
        nome_ativo=nome_ativo,
        data_base_str=data_base_start.strftime("%d/%m/%Y"),
        final_date=final_date,
        historico_normal=historico_normal,
        historico_punitivo=[{"data": d, "acumulado": val} for d, val in punitivo_mes.items()],
        valor_total_final=valor_total_ativo_final,
        ordem_cronologica=ordem_cronologica,
        valores_iniciais_str=valores_iniciais_str
    )

    # Nome final => "LMCalc_[nome]_DDMMAAAA.pdf"
    hoje_str = datetime.now().strftime("%d%m%Y")
    nome_sanitizado = nome_ativo.replace(" ", "_").replace("\"", "")
    nome_final = f"LMCalc_{nome_sanitizado}_{hoje_str}.pdf"

    return send_file(
        pdf_bytes,
        as_attachment=True,
        download_name=nome_final,
        mimetype="application/pdf"
    )

@app.route("/")
def home():
    return "API de Cálculo e PDF (sem data de cessão) - Online!"

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
