import streamlit as st
import pandas as pd
import re
from datetime import datetime
from rules import elabora_report, processa_budget, formatta_k_euro, MAPPATURA_DEFAULT, MESI_LISTA, MAP_MESI, colora_colonne_fornitori, genera_pdf_report

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