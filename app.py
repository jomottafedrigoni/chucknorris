import streamlit as st
import pandas as pd
import re
from datetime import datetime
import io
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

MESI_LISTA = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

MAP_MESI = {
    'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 
    'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08', 
    'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
}

MAPPATURA_DEFAULT = {
    "0085000PLA1": "Other",
    "0091200PLA1": "Canteen Expenses",
    "0083730DT": "Cleaning Expenses",
    "0083730PLA1": "Cleaning Expenses",
    "0079430DT": "Maintenance",
    "0079435PLA3": "Maintenance",
    "0079435PRO4": "Maintenance",
    "0090200PLA1": "Insurance",
    "0090210A1": "Insurance",
    "0083741PLA1": "Intercompany",
    "0120500A1": "Taxes and duties",
    "0083100PLA1": "Maintenance",
    "0083160PLA1": "Maintenance",
    "0083160PRO5": "Maintenance",
    "0083300PLA1": "Maintenance",
    "0091000PLA1": "Other",
    "0083700PLA1": "Consultancy",
    "0083700PLA3": "Consultancy",
    "0083710PLA1": "Consultancy",
    "0083710PRO4": "Consultancy",
    "0083720PLA1": "Consultancy",
    "0079412PLA1": "Spare Parts And Equipments",
    "0079420PRO4": "Spare Parts And Equipments",
    "0079412A1": "Spare Parts And Equipments",
    "0079420A1": "Spare Parts And Equipments",
    "0079410PLA1": "Cleaning Expenses",
    "0079410PRO1": "Cleaning Expenses",
    "0080640PLA1": "Spare Parts And Equipments",
    "0080640PRO1": "Spare Parts And Equipments",
    "0080640PRO5": "Spare Parts And Equipments",
    "0080650PLA1": "Spare Parts And Equipments",
    "0080650PRO5": "Spare Parts And Equipments",
    "0093600PLA1": "Security Costs",
    "0091100PLA1": "Training Costs",
    "0079400PLA1": "Travel Expenses",
    "0079400PLA3": "Travel Expenses",
    "0083750PLA1": "Travel Expenses",
    "0083750PLA2": "Travel Expenses",
    "0083750PLA3": "Travel Expenses",
    "0083750PLA4": "Travel Expenses",
    "0086000PLA1": "Travel Expenses",
    "0079412PRO4": "Spare Parts And Equipments",
    "0079412PRO5": "Spare Parts And Equipments",
    "0079420PLA1": "Spare Parts And Equipments",
    "0083100PRO5": "Maintenance",
    "0080650PRO1":	"Maintenance",
    "0083100PRO1":	"Maintenance"
}

def to_excel_download(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Budget_Elaborato')
    return output.getvalue()

def processa_budget(df_bdg, anno_ref=datetime.now().year):
    cols_mesi = [m for m in MESI_LISTA if m in df_bdg.columns]
    for m in cols_mesi:
        df_bdg[m] = pd.to_numeric(df_bdg[m], errors='coerce').fillna(0)
    
    df_melt = df_bdg.melt(
        id_vars=['Mappatura', 'Supplier'],
        value_vars=cols_mesi,
        var_name='Mese_Nome',
        value_name='Importo_BDG_Orig'
    )
    
    df_melt['Mese_Num'] = df_melt['Mese_Nome'].map(MAP_MESI)
    df_melt['Mese'] = str(anno_ref) + "-" + df_melt['Mese_Num']
    df_melt['Importo_BDG_kEUR'] = (df_melt['Importo_BDG_Orig'] * -1) / 1000.0
    
    df_melt['Mappatura'] = df_melt['Mappatura'].astype(str).str.strip().str.title()
    df_melt['Supplier'] = df_melt['Supplier'].astype(str).str.strip()
    
    return df_melt

def elabora_report(df_trans):
    df_out = df_trans.copy()
    df_out.columns = df_out.columns.astype(str).str.strip()

    col_desc_trovata = None
    nomi_possibili = [
        "Description"
    ]
    
    for col in nomi_possibili:
        if col in df_out.columns:
            col_desc_trovata = col
            break

    df_out["Descrizione"] = df_out[col_desc_trovata].fillna("-").astype(str) if col_desc_trovata else "-"

    for col in ["Importo nella valuta della transazione", "Importo", "Importo nella valuta di dichiarazione"]:
        if col in df_out.columns:
            df_out[col] = pd.to_numeric(df_out[col], errors='coerce').fillna(0)

    df_out["Mappatura"] = df_out["Account_Dim_Key"].map(MAPPATURA_DEFAULT)
    df_out = df_out.dropna(subset=["Mappatura"]).reset_index(drop=True)
    df_out = df_out.fillna("No Supplier")

    col_data = "Posting Date" if "Posting Date" in df_out.columns else ("Data documento" if "Data documento" in df_out.columns else None)
    col_imp = "Amount" if "Amount" in df_out.columns else None

    if col_data and col_imp:
        df_out["dt_temp"] = pd.to_datetime(df_out[col_data], dayfirst=True, errors='coerce')
        df_out["Anno"] = df_out["dt_temp"].dt.year.astype(str)
        df_out["Mese_Num"] = df_out["dt_temp"].dt.strftime('%m')
        df_out["Mese"] = df_out["Anno"] + "-" + df_out["Mese_Num"]
        df_out["Importo_ACT_kEUR"] = (df_out[col_imp] * -1) / 1000.0

    return df_out

def formatta_k_euro(val):
    if pd.isna(val) or round(val, 2) == 0:
        return "-"
    if val < 0:
        return f"({abs(val):,.2f})".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def colora_colonne_fornitori(df):
    if not isinstance(df, pd.DataFrame):
        return df

    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    for col in df.columns:
        col_str = str(col)
        is_tot_col = "TOTALE" in col_str.upper()

        if "ACT" in col_str:
            bg_color = '#d9d9d9' if is_tot_col else '#f2f2f2'
            base_style = f'background-color: {bg_color}; color: black;'
        elif "BDG" in col_str:
            bg_color = '#e6e6e6' if is_tot_col else '#ffffff'
            base_style = f'background-color: {bg_color}; color: black;'
        elif "Delta" in col_str:
            bg_color = '#b3d8ff' if is_tot_col else '#e6f2ff'
            base_style = f'background-color: {bg_color};'
        else:
            base_style = ''

        for idx in df.index:
            idx_str = str(idx).upper()
            is_tot_row = "TOTALE" in idx_str
            style = base_style

            if "Delta" in col_str:
                val = df.loc[idx, col]
                if isinstance(val, (int, float)):
                    if val > 0.01:
                        style += ' color: green;'
                    elif val < -0.01:
                        style += ' color: red;'
                    else:
                        style += ' color: black;'

            if is_tot_col or is_tot_row:
                style += ' font-weight: bold;'
                if is_tot_row and "Delta" not in col_str:
                    style += ' background-color: #d0d0d0;'
                elif is_tot_row and "Delta" in col_str:
                    style += ' background-color: #99ccff;'

            styles.loc[idx, col] = style

    return styles


def genera_pdf_report(df_act_f, df_bdg_f, df_bdg_raw, anno_sel, mese_ref):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=20,
        leftMargin=20,
        topMargin=25,
        bottomMargin=20
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('CoverTitle', parent=styles['Title'], fontSize=24, leading=28, alignment=1, textColor=colors.HexColor('#000000'))
    subtitle_style = ParagraphStyle('CoverSubtitle', parent=styles['Normal'], fontSize=11, leading=16, alignment=1, textColor=colors.HexColor('#444444'))
    disclaimer_style = ParagraphStyle('CoverDisclaimer', parent=styles['Normal'], fontSize=9, leading=13, alignment=1, textColor=colors.HexColor('#666666'))
    
    header_style = ParagraphStyle('PageHeader', parent=styles['Heading1'], fontSize=12, leading=15, textColor=colors.HexColor('#1A1A1A'), spaceAfter=8)
    sub_header_style = ParagraphStyle('SubHeader', parent=styles['Heading2'], fontSize=9, leading=11, textColor=colors.HexColor('#222222'), spaceBefore=6, spaceAfter=4)
    
    header_cell = ParagraphStyle('HCell', parent=styles['Normal'], fontSize=6.5, leading=7.5, fontName='Helvetica-Bold', textColor=colors.white, alignment=1)
    num_style = ParagraphStyle('NCell', parent=styles['Normal'], fontSize=6.5, leading=7.5, textColor=colors.black, alignment=2)
    num_bold = ParagraphStyle('NBold', parent=styles['Normal'], fontSize=6.5, leading=7.5, fontName='Helvetica-Bold', textColor=colors.black, alignment=2)
    txt_style = ParagraphStyle('TCell', parent=styles['Normal'], fontSize=6.5, leading=7.5, textColor=colors.black, alignment=0)
    txt_bold = ParagraphStyle('TBold', parent=styles['Normal'], fontSize=6.5, leading=7.5, fontName='Helvetica-Bold', textColor=colors.black, alignment=0)

    def formatta_contabile(val):
        if val == 0 or pd.isna(val):
            return "-"
        val_ass = abs(val)
        str_val = f"{val_ass:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"({str_val})" if val < 0 else str_val

    story = []

    # PAGINA COPERTINA
    story.append(Spacer(1, 40))
    story.append(Paragraph("Report Mensile Costi Fissi", title_style))
    story.append(Spacer(1, 10))
    story.append(Paragraph(f"Data Report: {datetime.now().strftime('%d/%m/%Y')}", subtitle_style))
    story.append(Spacer(1, 30))
    
    subtext = (
        "<b>Arconvert S.p.A. - Industrial Controlling Team</b><br/><br/>"
        "Il presente documento è rivolto unicamente al Plant Director di Arconvert S.p.A.<br/>"
        "Le informazioni contenute nel report sono strettamente confidenziali, riservate e destinate ad uso interno. "
        "Ne è severamente vietata la riproduzione, la diffusione o la condivisione con soggetti non autorizzati, "
        "sia all'interno che all'esterno dell'organizzazione.<br/><br/>"
        "Per chiarimenti sui dati contenuti nel report o richieste di approfondimento analitico, fare riferimento a:<br/>"
        "• <b>Industrial Controller Italy:</b> Jona Motta (jona.motta@fedrigoni.com)<br/>"
        "• <b>Group Industrial Controller:</b> Roberto Ghirardi (roberto.ghirardi@fedrigoni.com)"
    )
    story.append(Paragraph(subtext, disclaimer_style))
    story.append(PageBreak())

    def crea_tabella_actual(df_source, escludi_voci=[]):
        df_sub = df_source[~df_source["Mappatura"].isin(escludi_voci)].copy()
        col_group = "Mappatura"

        if df_sub.empty:
            return Paragraph("<i>Nessun dato Actual disponibile.</i>", txt_style)
            
        piv = df_sub.pivot_table(index="Mese", columns=col_group, values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
        piv["TOTALE MENSILE"] = piv.sum(axis=1)
        tot_map = piv.sum(axis=0)
        tot_map.name = "TOTALE"
        piv_full = pd.concat([piv, pd.DataFrame(tot_map).T])

        headers = ["Mese"] + [str(c) for c in piv_full.columns]
        table_data = [[Paragraph(h, header_cell) for h in headers]]

        for idx, row in piv_full.iterrows():
            is_tot = (str(idx) == "TOTALE")
            r_data = [Paragraph(str(idx), txt_bold if is_tot else txt_style)]
            for val in row:
                r_data.append(Paragraph(formatta_contabile(val), num_bold if is_tot else num_style))
            table_data.append(r_data)

        col_w = [45] + [(735 / (len(headers) - 1))] * (len(headers) - 1)
        t = Table(table_data, colWidths=col_w, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1A1A1A')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#CCCCCC')),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#F4F4F4')),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5),
            ('TOPPADDING', (0, 0), (-1, -1), 2.5),
        ]))
        return t

    def crea_tabella_act_vs_bdg(df_act_sub, df_bdg_sub, escludi_voci=[], is_ytd=False):
        delta_red = ParagraphStyle('DRed', parent=styles['Normal'], fontSize=6, leading=7, textColor=colors.HexColor('#C00000'), alignment=2)
        delta_red_bold = ParagraphStyle('DRedB', parent=styles['Normal'], fontSize=6, leading=7, fontName='Helvetica-Bold', textColor=colors.HexColor('#C00000'), alignment=2)
        delta_green = ParagraphStyle('DGreen', parent=styles['Normal'], fontSize=6, leading=7, textColor=colors.HexColor('#008000'), alignment=2)
        delta_green_bold = ParagraphStyle('DGreenB', parent=styles['Normal'], fontSize=6, leading=7, fontName='Helvetica-Bold', textColor=colors.HexColor('#008000'), alignment=2)
        
        num_s = ParagraphStyle('NCellS', parent=styles['Normal'], fontSize=6, leading=7, textColor=colors.black, alignment=2)
        num_b = ParagraphStyle('NBoldS', parent=styles['Normal'], fontSize=6, leading=7, fontName='Helvetica-Bold', textColor=colors.black, alignment=2)
        txt_s = ParagraphStyle('TCellS', parent=styles['Normal'], fontSize=6, leading=7, textColor=colors.black, alignment=0)
        txt_b = ParagraphStyle('TBoldS', parent=styles['Normal'], fontSize=6, leading=7, fontName='Helvetica-Bold', textColor=colors.black, alignment=0)

        def get_delta_style(v_delta, is_bold=False):
            if v_delta > -0.01:
                return delta_red_bold if is_bold else delta_red
            elif v_delta < 0.01:
                return delta_green_bold if is_bold else delta_green
            return num_b if is_bold else num_s

        df_act_cf = df_act_sub[~df_act_sub["Mappatura"].isin(escludi_voci)]
        df_bdg_cf = df_bdg_sub[~df_bdg_sub["Mappatura"].isin(escludi_voci)]

        piv_act = df_act_cf.pivot_table(index="Mese", columns="Mappatura", values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
        piv_bdg = df_bdg_cf.pivot_table(index="Mese", columns="Mappatura", values="Importo_BDG_kEUR", aggfunc="sum", fill_value=0)

        # Filtra i mesi fino a mese_ref se YTD, altrimenti mantiene tutti i 12 mesi per MTD
        if is_ytd and mese_ref:
            mesi_list = [f"{anno_sel}-{i:02d}" for i in range(1, 13) if f"{anno_sel}-{i:02d}" <= mese_ref]
        else:
            mesi_list = [f"{anno_sel}-{i:02d}" for i in range(1, 13)]

        if is_ytd:
            piv_act = piv_act.reindex(mesi_list, fill_value=0)
            piv_bdg = piv_bdg.reindex(mesi_list, fill_value=0)

            piv_act = piv_act.cumsum(axis=0)
            piv_bdg = piv_bdg.cumsum(axis=0)
        else:
            piv_act = piv_act.reindex(mesi_list, fill_value=0)
            piv_bdg = piv_bdg.reindex(mesi_list, fill_value=0)

        cats = sorted(list(set(piv_act.columns).union(set(piv_bdg.columns))))
        if not cats:
            return Paragraph("<i>Nessun dato Costi Fissi disponibile.</i>", txt_s)

        h_row1 = [Paragraph("M", header_cell)]
        for c in cats:
            h_row1.extend([Paragraph(c, header_cell), "", ""])
        h_row1.extend([Paragraph("TOT", header_cell), "", ""])

        h_row2 = [""]
        for _ in range(len(cats) + 1):
            h_row2.extend([Paragraph("ACT", header_cell), Paragraph("BDG", header_cell), Paragraph("&Delta;", header_cell)])

        table_data = [h_row1, h_row2]

        tot_act_cats = {c: 0.0 for c in cats}
        tot_bdg_cats = {c: 0.0 for c in cats}
        tot_act_gen, tot_bdg_gen = 0.0, 0.0

        riga_mese_chiuso_idx = None

        for m_str in mesi_list:
            m_breve = m_str[5:] if len(m_str) >= 7 else m_str
            row = [Paragraph(m_breve, txt_s)]
            m_act_tot, m_bdg_tot = 0.0, 0.0

            # Salva l'indice di riga se corrisponde al mese di riferimento/chiuso
            if m_str == mese_ref:
                riga_mese_chiuso_idx = len(table_data)

            for c in cats:
                v_bdg = piv_bdg.loc[m_str, c] if (m_str in piv_bdg.index and c in piv_bdg.columns) else 0.0
                tot_bdg_cats[c] += v_bdg
                m_bdg_tot += v_bdg

                v_act = piv_act.loc[m_str, c] if (m_str in piv_act.index and c in piv_act.columns) else 0.0
                v_delta = v_bdg - v_act
                tot_act_cats[c] += v_act
                m_act_tot += v_act

                row.extend([
                    Paragraph(formatta_contabile(v_act), num_s),
                    Paragraph(formatta_contabile(v_bdg), num_s),
                    Paragraph(formatta_contabile(v_delta), get_delta_style(v_delta))
                ])

            m_delta_tot = m_bdg_tot - m_act_tot
            tot_act_gen += m_act_tot
            tot_bdg_gen += m_bdg_tot

            row.extend([
                Paragraph(formatta_contabile(m_act_tot), num_b),
                Paragraph(formatta_contabile(m_bdg_tot), num_b),
                Paragraph(formatta_contabile(m_delta_tot), get_delta_style(m_delta_tot, is_bold=True))
            ])
            table_data.append(row)

        if not is_ytd:
            tot_row = [Paragraph("TOT", txt_b)]
            for c in cats:
                v_act = tot_act_cats[c]
                v_bdg = tot_bdg_cats[c]
                v_delta = v_bdg - v_act
                tot_row.extend([
                    Paragraph(formatta_contabile(v_act), num_b),
                    Paragraph(formatta_contabile(v_bdg), num_b),
                    Paragraph(formatta_contabile(v_delta), get_delta_style(v_delta, is_bold=True))
                ])

            tot_delta_gen = tot_bdg_gen - tot_act_gen
            tot_row.extend([
                Paragraph(formatta_contabile(tot_act_gen), num_b),
                Paragraph(formatta_contabile(tot_bdg_gen), num_b),
                Paragraph(formatta_contabile(tot_delta_gen), get_delta_style(tot_delta_gen, is_bold=True))
            ])
            table_data.append(tot_row)

        num_cols = 1 + (len(cats) + 1) * 3
        col_w = [20] + [(783 / (num_cols - 1))] * (num_cols - 1)

        t_style = [
            ('BACKGROUND', (0, 0), (-1, 1), colors.HexColor('#1A1A1A')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#CCCCCC')),
            ('SPAN', (0, 0), (0, 1)),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1.5),
            ('TOPPADDING', (0, 0), (-1, -1), 1.5),
        ]

        # Evidenzia in grigio chiaro la riga del mese chiuso (nella tabella MTD)
        if not is_ytd and riga_mese_chiuso_idx is not None:
            t_style.append(('BACKGROUND', (0, riga_mese_chiuso_idx), (-1, riga_mese_chiuso_idx), colors.HexColor('#E0E0E0')))

        if not is_ytd:
            t_style.append(('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#F4F4F4')))

        col_idx = 1
        for _ in range(len(cats) + 1):
            t_style.append(('SPAN', (col_idx, 0), (col_idx + 2, 0)))
            col_idx += 3

        t = Table(table_data, colWidths=col_w, repeatRows=2)
        t.setStyle(TableStyle(t_style))
        return t

    voci_fisse_escluse_act = []
    voci_fisse_escluse_bdg = voci_fisse_escluse_act + ["Insurance", "Intercompany","Spare Parts and Equipments", "Subscription And Associations Fees","Taxes and duties","Training Costs"]

    sezioni = [
        ("Arconvert S.p.A.", df_act_f, df_bdg_f)
    ]

    for i, (titolo, df_act_sub, df_bdg_sub) in enumerate(sezioni):
        story.append(Paragraph(f"<b>{titolo} - Costi Fissi</b>", header_style))
        story.append(Paragraph("1. Costi Fissi - Actual MTD (in k€)", sub_header_style))
        story.append(crea_tabella_actual(df_act_sub, escludi_voci=voci_fisse_escluse_act))
        story.append(Spacer(1, 6))
        story.append(Paragraph("2. Costi Fissi - Actual vs Budget MTD (in k€)", sub_header_style))
        story.append(crea_tabella_act_vs_bdg(df_act_sub, df_bdg_sub, escludi_voci=voci_fisse_escluse_bdg, is_ytd=False))
        story.append(Spacer(1, 6))
        story.append(Paragraph("3. Costi Fissi - Actual vs Budget YTD (in k€)", sub_header_style))
        story.append(crea_tabella_act_vs_bdg(df_act_sub, df_bdg_sub, escludi_voci=voci_fisse_escluse_bdg, is_ytd=True))
        story.append(PageBreak())

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()




st.set_page_config(
    page_title="Costi Fissi Arconvert Spa",
    page_icon=":central_african_republic:",
    layout="wide"
)



st.title("ARCONVERT Report Costi Fissi")

st.logo("logo.png")
st.markdown(
    """
    <style>
        [data-testid="stSidebarHeader"] img {
            height: 100px !important;
            width: auto !important;
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.sidebar.header("Caricamento File")
uploaded_transazioni = st.sidebar.file_uploader("1. File Transazioni (ACT)", type=["xlsx", "xls"], key="transazioni")
uploaded_budget = st.sidebar.file_uploader("3. File Budget (BDG)", type=["xlsx", "xls"], key="budget")

files_caricati = uploaded_transazioni and uploaded_budget

if not files_caricati:
    pattern_7_cifre = re.compile(r"\d{7}")
    conti_list = []

    for chiave in MAPPATURA_DEFAULT.keys():
        # Rimuove eventuali punti dalla chiave per contare solo le cifre numeriche
        chiave_pulita = str(chiave).strip().replace(".", "")
        match = pattern_7_cifre.search(chiave_pulita)

        if match:
            conti_list.append(match.group(0))

    conti_uniche = list(dict.fromkeys(conti_list))

    st.info("Carica i file richiesti nella barra laterale per avviare l'analisi.")
    st.markdown("---")
    st.subheader("Istruzioni per l'estrazione dei file sorgente")

    with st.expander("Istruzioni per il funzionamento del report", expanded=True):
        with st.expander("Estrazione File Transazioni (Actual)", expanded=True):
            st.markdown(
                """
            **Procedura di estrazione:**
            Jona, fatti dare un service account altrimenti sono i tuoi parametri che vengono usati e questo non va bene per l'azienda.
            """
            )

            st.write("**Lista Conti inclusi nella Mappatura Default:**")
            testo_conti = "\n".join(conti_uniche)
            st.code(testo_conti, language="text")

        with st.expander("Estrazione file Anagrafica Fornitori", expanded=True):
            st.markdown(
                """
            Jona, pure qui, pure qui.
            """
            )

        with st.expander("Estrazione file di Budget", expanded=True):
            st.markdown(
                "📌 **Nota:** Richiedere il file di budget al team di **Industrial Controlling**."
            )

if uploaded_transazioni and uploaded_budget:
    try:
        df_act_raw = pd.read_excel(uploaded_transazioni)
        df_bdg_raw = pd.read_excel(uploaded_budget)
        df_act = elabora_report(df_act_raw)
        st.sidebar.markdown("---")
        anni_disponibili = sorted([a for a in df_act["Anno"].unique() if a != "nan"], reverse=True)
        anno_sel = st.sidebar.selectbox("Seleziona Anno:", anni_disponibili, index=0)
        oggi = datetime.now()
        mese_prec_num = 12 if oggi.month == 1 else oggi.month - 1
        mese_chiuso_ref = f"{anno_sel}-{mese_prec_num:02d}"
        df_bdg = processa_budget(df_bdg_raw, anno_ref=int(anno_sel) if anno_sel else 2026)
        mappature_totali = sorted(list(set(df_act["Mappatura"].unique()).union(set(df_bdg["Mappatura"].unique()))))
        mappatura_sel = st.sidebar.multiselect("Seleziona Voce di Spesa:", options=mappature_totali, default=[])
        mesi_disponibili = sorted(list(set(df_act["Mese"].unique()).union(set(df_bdg["Mese"].unique()))))
        mesi_sel = st.sidebar.multiselect("Seleziona Mesi:", options=mesi_disponibili, default=[])
        df_act_f = df_act[df_act["Anno"] == anno_sel].copy()
        df_bdg_f = df_bdg.copy()
        if mappatura_sel:
                df_act_f = df_act_f[df_act_f["Mappatura"].isin(mappatura_sel)]
                df_bdg_f = df_bdg_f[df_bdg_f["Mappatura"].isin(mappatura_sel)]
        if mesi_sel:
                df_act_f_mesi = df_act_f[df_act_f["Mese"].isin(mesi_sel)]
        else:
            df_act_f_mesi = df_act_f.copy()

        tab_dettaglio, tab_act_bdg = st.tabs(["Report Analisi Spese", "ACT vs BDG"])

        with tab_dettaglio:
            df_sezione = df_act_f_mesi.copy()
            col_raggruppamento = "Mappatura"
            if mesi_sel:
                m_max = max(mesi_sel)
            else:
                mesi_act_validi = [m for m in df_act_f["Mese"].unique() if pd.notna(m) and m <= mese_chiuso_ref]
                m_max = max(mesi_act_validi) if mesi_act_validi else (max(df_act_f["Mese"].unique()) if len(df_act_f) > 0 else "")
            pdf_bytes = genera_pdf_report(df_act_f, df_bdg_f, df_bdg_raw, anno_sel, mese_chiuso_ref)
            st.sidebar.markdown("---")
            st.sidebar.download_button(
                label="Scarica Report PDF",
                data=pdf_bytes,
                file_name=f"Report Costi Fissi {mese_chiuso_ref}.pdf",
                mime="application/pdf"
            )
            st.header("Fixed Costs - Dettaglio Analisi Spese")
            #quiquiquipdf creation
            s_mtd = df_sezione[df_sezione["Mese"] == m_max]["Importo_ACT_kEUR"].sum() if m_max else 0
            s_ytd = df_sezione[df_sezione["Mese"] <= m_max]["Importo_ACT_kEUR"].sum() if m_max else 0
            st.markdown("---")
            st.subheader(f"Actual MTD per Voci di Spesa (in k€)")
            piv_mensile = df_sezione.pivot_table(index="Mese", columns=col_raggruppamento, values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
            piv_mensile["TOTALE MENSILE"] = piv_mensile.sum(axis=1)
            tot_map = piv_mensile.sum(axis=0)
            tot_map.name = "TOTALE"
            piv_mensile_full = pd.concat([piv_mensile, pd.DataFrame(tot_map).T])
            st.dataframe(piv_mensile_full.style.format(formatta_k_euro), use_container_width=True)
            st.subheader("Actual MTD per Plant (k€)")
            piv_plant = df_sezione.pivot_table(index="Mese", values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
            st.bar_chart(piv_plant)
            st.markdown("---")
            st.subheader(f"Actual YTD per Fornitore e Voci di Spesa (in k€)")
            piv_ytd = df_sezione.pivot_table(index="Supplier", columns=col_raggruppamento, values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
            piv_ytd["TOTALE FORNITORE"] = piv_ytd.sum(axis=1)
            piv_ytd = piv_ytd.sort_values(by="TOTALE FORNITORE", ascending=False)
            tot_cat = piv_ytd.sum(axis=0)
            tot_cat.name = "TOTALE CATEGORIA"
            piv_ytd_full = pd.concat([piv_ytd, pd.DataFrame(tot_cat).T])
            st.dataframe(piv_ytd_full.style.format(formatta_k_euro), use_container_width=True)

            st.markdown("---")
            st.subheader("Dettaglio Analitico Transazioni per Fornitore")
            
            cols_dettaglio = [c for c in ["Scenario", "Posting Date", "G_L Account No_", "Mappatura", "Supplier", "Importo_ACT_kEUR", "Descrizione"] if c in df_sezione.columns]
            df_dett = df_sezione[cols_dettaglio].copy()
            if "Importo_ACT_kEUR" in df_dett.columns:
                df_dett = df_dett.rename(columns={"Importo_ACT_kEUR": "Importo (k€)"})
            
            st.dataframe(
                df_dett.style.format({"Importo (k€)": formatta_k_euro}),
                use_container_width=True
            )

        with tab_act_bdg:
            tipo_vista = st.radio("Orizzonte Temporale:", ["MTD (Valori Mensili Puntuali)", "YTD (Cumulato Progressivo Mese per Mese)"], horizontal=True)
            st.header("Fixed Costs - ACT vs BDG")
            st.subheader("MTD / YTD Actual vs Budget per Voce di Spesa")
            df_act_m1=df_act_f.copy()
            df_bdg_m1=df_bdg_f.copy()
            piv_act_m = df_act_m1.pivot_table(index="Mese", columns="Mappatura", values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
            piv_bdg_m = df_bdg_m1.pivot_table(index="Mese", columns="Mappatura", values="Importo_BDG_kEUR", aggfunc="sum", fill_value=0)
            mesi_completi = [f"{anno_sel}-{idx:02d}" for idx in range(1, 13)]
            if "YTD" in tipo_vista:
                mesi_con_dati = sorted([m for m in piv_act_m.index if m in mesi_completi and (piv_act_m.loc[m] != 0).any()])
                ultimo_mese_act = max(mesi_con_dati) if mesi_con_dati else None

                piv_act_m = piv_act_m.reindex(mesi_completi)
                piv_bdg_m = piv_bdg_m.reindex(mesi_completi, fill_value=0)

                piv_bdg_m = piv_bdg_m.cumsum(axis=0)
                piv_act_m = piv_act_m.cumsum(axis=0)

                if ultimo_mese_act:
                    mesi_successivi = [m for m in piv_act_m.index if m > ultimo_mese_act]
                    for m in mesi_successivi:
                        piv_act_m.loc[m] = piv_act_m.loc[ultimo_mese_act]
            else:
                piv_act_m = piv_act_m.reindex(mesi_completi, fill_value=0)
                piv_bdg_m = piv_bdg_m.reindex(mesi_completi, fill_value=0)
            all_maps = sorted(list(set(piv_act_m.columns).union(set(piv_bdg_m.columns))))
                
            rows_m1 = []
            for idx, m_code in enumerate(MESI_LISTA, 1):
                m_str = f"{anno_sel}-{idx:02d}"
                r = {"Mese": f"{m_code} ({m_str})"}
                tot_act_mese = 0.0
                tot_bdg_mese = 0.0

                for m_cat in all_maps:
                    v_act = piv_act_m.loc[m_str, m_cat] if (m_str in piv_act_m.index and m_cat in piv_act_m.columns and pd.notna(piv_act_m.loc[m_str, m_cat])) else 0.0
                    v_bdg = piv_bdg_m.loc[m_str, m_cat] if (m_str in piv_bdg_m.index and m_cat in piv_bdg_m.columns and pd.notna(piv_bdg_m.loc[m_str, m_cat])) else 0.0
                    
                    r[f"{m_cat} ACT"] = v_act
                    r[f"{m_cat} BDG"] = v_bdg
                    r[f"{m_cat} Delta"] = v_act - v_bdg

                    tot_act_mese += v_act
                    tot_bdg_mese += v_bdg

                r["TOTALE ACT"] = tot_act_mese
                r["TOTALE BDG"] = tot_bdg_mese
                r["TOTALE Delta"] = tot_act_mese - tot_bdg_mese

                rows_m1.append(r)

            df_mod1 = pd.DataFrame(rows_m1).set_index("Mese")

            if "YTD" not in tipo_vista:
                tot_dict = {}
                for col in df_mod1.columns:
                    if "Delta" in col:
                        base_col = col.replace(" Delta", "")
                        tot_dict[col] = tot_dict[f"{base_col} ACT"] - tot_dict[f"{base_col} BDG"]
                    else:
                        tot_dict[col] = df_mod1[col].sum()
                
                df_tot_m1 = pd.DataFrame([tot_dict], index=["TOTALE ANNO"])
                df_mod1_full = pd.concat([df_mod1, df_tot_m1])
            else:
                df_mod1_full = df_mod1

            st.dataframe(
                df_mod1_full.style.format(formatta_k_euro).apply(colora_colonne_fornitori, axis=None),
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Errore generale nell'elaborazione dei file: {e}")