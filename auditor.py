import os
import pandas as pd
import pdfplumber
from collections import defaultdict
import re
from typing import Tuple

def identify_spacer_note(text):
    for line in text.split("\n"):
        if "breaker space - breaker is not included" in line.lower():
            note_number = line.strip().split()[0]
            return note_number.replace(".", "")
    return None

def extract_breakers_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page_count = len(pdf.pages)
        breakers_text = ""
        doble_fv = False

        if page_count == 2:
            breakers_text = pdf.pages[1].extract_text() or ""
        elif page_count == 3:
            breakers_text = (pdf.pages[1].extract_text() or "") + "\n" + (pdf.pages[2].extract_text() or "")
        else:
            panel_marks_pages = []
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                if "PANEL MARKS" in text:
                    panel_marks_pages.append(i)

            if len(panel_marks_pages) >= 2:
                doble_fv = True
                segunda_fv_index = panel_marks_pages[1]
                lectura_index = segunda_fv_index + 1 if segunda_fv_index + 1 < page_count else segunda_fv_index
                breakers_text = pdf.pages[lectura_index].extract_text() or ""
            else:
                breakers_text = pdf.pages[1].extract_text() or ""

        spacer_note = identify_spacer_note(breakers_text)
        lines = breakers_text.split("\n")
        breakers = []
        for line in lines:
            parts = line.strip().split()
            note_detected = parts[-1] if len(parts) > 1 else ""
            for part in parts:
                if (part.startswith("XT") or part.startswith("TEY")) and len(part) >= 10 and "SPACE" not in part.upper():
                    if spacer_note and spacer_note in note_detected:
                        continue
                    breakers.append(part.strip("-.,"))

        breaker_counts = defaultdict(int)
        for b in breakers:
            breaker_counts[b] += 1

        sap_order = os.path.splitext(os.path.basename(pdf_path))[0]
        sap_order = re.sub(r" \(\d+\)$", "", sap_order)

        extracted = []
        for np_alt, qty in breaker_counts.items():
            extracted.append({
                "SAP Order": sap_order,
                "NP Alternativo": np_alt,
                "Cantidad": qty,
                "Doble FV": doble_fv
            })

        return extracted

def ejecutar_auditoria(audit_folder: str) -> Tuple[str, str]:
    input_folder = os.path.join(audit_folder, "Información Auditada", "FVs Auditados")
    equivalencias_file = os.path.join(os.path.dirname(__file__), "np_equivalencias.xlsx")
    info_auditada = os.path.join(audit_folder, "Información Auditada")

    cruce_file = next((f for f in os.listdir(info_auditada) if f.upper().startswith("CRUCE_") and f.lower().endswith(".xlsx")), None)
    bom_file = next((f for f in os.listdir(info_auditada) if f.upper().startswith("BOM_") and f.lower().endswith(".xlsx")), None)

    cruce_file = os.path.join(info_auditada, cruce_file)
    bom_file = os.path.join(info_auditada, bom_file)

    for path in [equivalencias_file, cruce_file, bom_file, input_folder]:
        if not os.path.exists(path):
            raise FileNotFoundError(f"No se encontró el archivo o carpeta: {path}")

    df_equiv = pd.read_excel(equivalencias_file)
    df_equiv.rename(columns={df_equiv.columns[0]: "NP ABB", df_equiv.columns[1]: "NP Alternativo"}, inplace=True)
    df_equiv["NP ABB"] = df_equiv["NP ABB"].astype(str).str.strip().str.upper()
    df_equiv["NP Alternativo"] = df_equiv["NP Alternativo"].astype(str).str.strip().str.upper()

    df_cruce = pd.read_excel(cruce_file)
    df_cruce.columns = [col.strip() for col in df_cruce.columns]
    df_cruce["Sales Order"] = df_cruce["Sales Order"].astype(str).str.replace(".0", "").str.strip()
    df_cruce["Sales order item"] = df_cruce["Sales order item"].astype(str).str.replace(".0", "").str.strip()
    df_cruce["Concatenado"] = df_cruce["Sales Order"] + "-" + df_cruce["Sales order item"]

    results = []
    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(input_folder, filename)
            extracted = extract_breakers_from_pdf(pdf_path)
            results.extend(extracted)

    print("BREAKERS EXTRAÍDOS:", results)


    df_breakers = pd.DataFrame(results)
    df_breakers["Concatenado"] = df_breakers["SAP Order"].astype(str)
    df_breakers = df_breakers.merge(df_equiv, on="NP Alternativo", how="left")
    df_breakers = df_breakers.merge(
        df_cruce[["Concatenado", "Order"]].rename(columns={"Order": "PO"}),
        on="Concatenado", how="left"
    )
    df_breakers["PO"] = df_breakers["PO"].apply(lambda x: str(int(x)) if pd.notna(x) else "").str.strip()
    df_breakers["NP ABB"] = df_breakers["NP ABB"].astype(str).str.strip().str.upper()

    df_breakers_final = df_breakers[["PO", "NP Alternativo", "NP ABB", "SAP Order", "Cantidad"]]
    df_breakers_final = df_breakers_final[df_breakers_final["NP ABB"].notnull()]
    output1 = os.path.join(audit_folder, "breakers_por_orden.xlsx")
    df_breakers_final.to_excel(output1, index=False)

    df_bom = pd.read_excel(bom_file)
    df_bom.columns = [col.strip() for col in df_bom.columns]
    df_bom.rename(columns={
        df_bom.columns[0]: "PO",
        df_bom.columns[1]: "NP ABB",
        df_bom.columns[4]: "Cantidad SAP",
        df_bom.columns[5]: "BOM Item"
    }, inplace=True)
    df_bom["PO"] = df_bom["PO"].apply(lambda x: str(int(x)) if pd.notna(x) else "").str.strip()
    df_bom["NP ABB"] = df_bom["NP ABB"].astype(str).str.strip().str.upper()
    df_bom = df_bom[df_bom["NP ABB"].str.startswith("1SDX") | df_bom["NP ABB"].isin(df_equiv["NP ABB"])]
    df_bom["Es Hijo"] = ~df_bom["BOM Item"].astype(str).str.startswith("B")
    df_bom["Comentario"] = df_bom["Es Hijo"].apply(lambda x: "Item hijo" if x else "Item padre")

    resultados = []
    for _, row in df_breakers.iterrows():
        po = row["PO"]
        np_abb = row["NP ABB"]
        cantidad_fv = row["Cantidad"]
        concatenado = row["SAP Order"]
        np_alternativo = row["NP Alternativo"]
        doble_fv_flag = row.get("Doble FV", False)

        bom_match = df_bom[(df_bom["PO"] == po) & (df_bom["NP ABB"] == np_abb)]

        if not bom_match.empty:
            cantidad_bom = bom_match["Cantidad SAP"].sum()
            estado = "✅ Correcto" if cantidad_bom == cantidad_fv else "❌ Diferencia"
            comentario = ""
        else:
            cantidad_bom = 0
            estado = "❌ Faltante"
            comentario = "No encontrado en BOM"

        if doble_fv_flag:
            comentario = "Doble FV"

        resultados.append({
            "Concatenado": concatenado,
            "PO": po,
            "Alternativo": np_alternativo,
            "NP ABB": np_abb,
            "FV Qty": cantidad_fv,
            "BOM Qty": cantidad_bom,
            "Estado": estado,
            "Comentario": comentario
        })

    df_resultado = pd.DataFrame(resultados)
    output2 = os.path.join(audit_folder, "verificacion_completa.xlsx")
    df_resultado.to_excel(output2, index=False)

    return output1, output2
