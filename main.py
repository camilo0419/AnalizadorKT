import os
import re
import threading
import logging
import unicodedata
import pandas as pd
from tkinter import filedialog, Tk, Button, Label, messagebox, ttk, Entry, Toplevel, StringVar
from pdfminer.high_level import extract_text
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# =========================
# Configuración general
# =========================
MRC_BASE = "base_clausulas_mrc.xlsx"
TTE_BASE = "base_clausulas_transporte.xlsx"

for name in ["pdfminer", "pdfminer.layout", "pdfminer.converter", "pdfminer.image", "pdfminer.pdfinterp"]:
    logging.getLogger(name).setLevel(logging.ERROR)

# =========================
# Normalización y utilidades
# =========================
IGNORE_TOKENS = {
    "nro", "numero", "no", "opcionales", "opcional", "plan",
    "sub", "sublimite", "sublimites", "sub-limite", "sub-limites",
    "lote", "anexo", "anexos"
}

TITLE_ANCHORS = (
    "amparo|clausula|cláusula|clausulado|asistencia|lucro|sustraccion|sustracción|equipo|equipos|maquinaria|"
    "terremoto|asonada|responsabilidad|rotura|definicion|definición|obras|arte|reposicion|reposición|"
    "archivo|archivos|hurto|incendio|explosion|explosión|robo|electric|eléctric|averia|avería|terceros|dano|"
    "gasto|gastos|arrendamiento|incremento|condiciones|generales|limite|límite|"
    "modulo|extension|ampliacion|caravana|merma|contaminacion|experticio|aviso|bienes|prima|minima|minimo|"
    "marcas|fabrica|averia particular|carga|descarga|transito|viaje|itinerario|"
    "valores|mercancia|mercancias|contenedores|remocion|escombros|refrigeracion|calefaccion|"
    "ferias|exposiciones|multimodal|devoluciones|redespachos|compensacion|liberacion|salvamento|ajustador|"
    "exclusion|exclusiones|nacionalizacion|interes|contingente|horario|declaracion"
)

STOP_WORDS = {
    "de","del","la","el","los","las","y","o","por","para","en","a","con","sin",
    "amparo","clausula","clausulas","cláusula","cláusulas","nro","opcional","opcion","opciones"
}

PAGE_BREAK = "\f"

CANON_HEADERS = [
    "No.",
    "Cláusula",
    "Texto de la cláusula",
    "Tipo de operación",
    "Valor asegurado",
    "Observaciones."
]

def strip_accents(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def _normalize(s: str) -> str:
    s = strip_accents(str(s)).lower()
    s = re.sub(r'\s+', ' ', s).strip()
    s = s.replace("operación", "operacion").replace("asegurado", "asegurado")
    return s

HEADER_SYNONYMS = {
    "No.": ["no", "no.", "n°", "nro", "numero", "n.º", "nº"],
    "Cláusula": ["clausula", "cláusula", "multiriesgo corporativo"],
    "Texto de la cláusula": ["texto de la clausula", "texto de la cláusula", "texto clausula", "texto de clausula"],
    "Tipo de operación": ["tipo de operacion", "tipo de operación", "tipo de\noperación", "tipo operacion"],
    "Valor asegurado": ["valor asegurado", "suma asegurada", "valor a segurar"],
    "Observaciones.": ["observaciones.", "observaciones", "observaciones "],
}

def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    norm_cols = {_normalize(c): c for c in df.columns}
    for target, variants in HEADER_SYNONYMS.items():
        for v in variants:
            nv = _normalize(v)
            for col_norm, original in norm_cols.items():
                if col_norm == nv:
                    mapping[original] = target
                    break
            if target in mapping.values():
                break
    if mapping:
        df = df.rename(columns=mapping)
    # Asegura que existan todas las columnas canónicas (crea vacías si faltan)
    for col in CANON_HEADERS:
        if col not in df.columns:
            df[col] = ""
    # Reordena
    df = df[CANON_HEADERS]
    return df

def normalize_spaces_keep_newlines(s: str) -> str:
    s = re.sub(r'[ \t\r\f\v]+', ' ', s); s = re.sub(r' ?\n ?', '\n', s); return s

def fix_hyphen_linebreaks(s: str) -> str:
    return re.sub(r'-\s*\n\s*', '', s)

def fix_spaced_caps(s: str) -> str:
    def join_spaced_caps(m): return m.group(0).replace(' ', '')
    pattern = r'(?:(?<=\s)|^)(?:[A-ZÁÉÍÓÚÑ]\s){3,}[A-ZÁÉÍÓÚÑ](?=\s|[,.;:()\-\n]|$)'
    return re.sub(pattern, join_spaced_caps, s)

def repair_upper_sequences(pdf_text: str) -> str:
    parts = pdf_text.split(PAGE_BREAK)
    repaired_pages = []
    for part in parts:
        lines = part.splitlines(); out, i = [], 0
        while i < len(lines):
            line = lines[i]
            if line.strip().isupper() and 1 <= len(line.strip()) <= 32:
                j, seq = i, []
                while j < len(lines) and (lines[j].strip().isupper() or lines[j].strip() == '') and len(seq) < 6:
                    if lines[j].strip(): seq.append(lines[j].strip())
                    j += 1
                if len(seq) >= 2: out.append(' '.join(seq)); i = j; continue
            out.append(line); i += 1
        repaired_pages.append('\n'.join(out))
    return PAGE_BREAK.join(repaired_pages)

def normalize_pdf_text(pdf_path: str):
    raw = extract_text(pdf_path)
    step1 = fix_hyphen_linebreaks(raw)
    step2 = repair_upper_sequences(step1)
    step3 = fix_spaced_caps(step2)
    repaired_text = step3
    norm_lines = strip_accents(repaired_text.lower())
    norm_lines = normalize_spaces_keep_newlines(norm_lines)
    norm_compact = re.sub(r'\s+', ' ', norm_lines.replace(PAGE_BREAK, ' ')).strip()
    return repaired_text, norm_lines, norm_compact

def _cut_after_line(text: str, pos_end: int) -> str:
    line_end = text.find('\n', pos_end)
    if line_end == -1: line_end = len(text)
    cut_pos = min(line_end + 1, len(text))
    return text[cut_pos:]

def crop_from_cp_and_clausulas(repaired_text: str, *, cortar_clausulas: bool = True):
    raw = repaired_text; low = strip_accents(raw.lower())
    m_cp = re.search(r'^\s*condiciones\s+particulares\s*:?\s*$', low, flags=re.M)
    if m_cp: raw = _cut_after_line(raw, m_cp.end()); low = strip_accents(raw.lower())
    if cortar_clausulas:
        m_cl = re.search(r'^\s*cl[aá]usulas\s*:?\s*$', low, flags=re.M)
        if m_cl: raw = _cut_after_line(raw, m_cl.end())
    norm_lines = strip_accents(raw.lower()); norm_lines = normalize_spaces_keep_newlines(norm_lines)
    norm_compact = re.sub(r'\s+', ' ', norm_lines.replace(PAGE_BREAK, ' ')).strip()
    return raw, norm_lines, norm_compact

def normalized_plain_strict(s: str) -> str:
    s = strip_accents(s.lower()); s = re.sub(r'[^a-z0-9]+', ' ', s); return re.sub(r'\s+', ' ', s).strip()

def singularize_token(tok: str) -> str:
    if len(tok) > 4 and tok.endswith('s'): return tok[:-1]
    return tok

def normalized_plain_canonical(s: str) -> str:
    s = strip_accents(s.lower()); s = re.sub(r'[/-]', ' ', s); s = re.sub(r'[^a-z0-9 ]+', ' ', s)
    toks = [t for t in s.split() if t and t not in IGNORE_TOKENS]
    return ' '.join(singularize_token(t) for t in toks).strip()

def has_title_anchor(s: str) -> bool:
    return re.search(r'\b(?:' + TITLE_ANCHORS + r')\b', s) is not None

def is_body_line(s: str) -> bool:
    T = s.strip().lower()
    if re.match(r'^\s*\d+\)\s', T): return True
    if re.search(r'[$€%]|vigencia|por\s+evento|deducible|prima', T): return True
    if T.endswith('.') and len(T.split()) >= 10: return True
    return False

def is_title_like(norm_text: str, orig_text: str) -> bool:
    T = norm_text.strip()
    if len(T) < 5 or len(T) > 320: return False
    if T.count('.') > 8: return False
    if re.match(r'^\d+\)\s', T): return False
    if re.search(r'[$€]|vigencia|por\s+evento|prima|deducible', T): return False
    if T.endswith('.') and len(T.split()) >= 12: return False
    if re.search(r'(^|\s)\d+\.\s', T): return True
    if has_title_anchor(T): return True
    if (len(T.split()) <= 12) and (not orig_text.strip().endswith('.')) and (not is_body_line(orig_text)):
        return True
    return False

def build_title_candidates(norm_lines: str, repaired_text: str, max_lines_combo=6):
    rep_lines = repaired_text.split('\n'); norm_lns = norm_lines.split('\n')
    assert len(rep_lines) == len(norm_lns)
    line_starts, acc = [], 0
    for s in rep_lines:
        line_starts.append(acc); acc += len(s) + 1
    cand = []; N = len(rep_lines)
    for i in range(N):
        if not norm_lns[i].strip(): continue
        if is_body_line(rep_lines[i]): continue
        for k in range(1, max_lines_combo + 1):
            j = i + k - 1
            if j >= N: break
            if k > 1 and any(is_body_line(rep_lines[t]) for t in range(i + 1, j + 1)): break
            combo_norm = ' '.join(norm_lns[i:j+1]).strip()
            combo_orig = ' '.join(rep_lines[i:j+1]).strip()
            if not is_title_like(combo_norm, combo_orig): continue
            if combo_orig.strip().endswith('.') and len(combo_orig.split()) >= 12: continue
            cand.append({
                'text_strict': normalized_plain_strict(combo_norm),
                'text_canon' : normalized_plain_canonical(combo_norm),
                'pos'        : line_starts[i],
                'orig'       : re.sub(r'\s+', ' ', combo_orig)
            })
    seen, uniq = set(), []
    for c in cand:
        key = (c['text_strict'], c['pos'])
        if key not in seen:
            seen.add(key); uniq.append(c)
    uniq.sort(key=lambda x: x['pos'])
    return uniq

def jaccard(a: set, b: set) -> float:
    if not a or not b: return 0.0
    inter = len(a & b); union = len(a | b)
    return inter / union if union else 0.0

# -------- Parámetros por ramo (automático, sin UI) --------
def get_params(excel_base: str):
    is_tte = "transporte" in os.path.basename(excel_base).lower()
    if is_tte:
        # Transporte: flexible (recall alto)
        return dict(
            MIN_LEN=6, JACCARD=0.78, REQUIRE_ANCHOR=False, REQUIRE_FIRST_TOKEN=False,
            MAX_LINES=6, FOOTER_K=12, CORTAR_CLAUSULAS=False
        )
    # MRC: más estricto (precisión)
    return dict(
        MIN_LEN=10, JACCARD=0.88, REQUIRE_ANCHOR=True, REQUIRE_FIRST_TOKEN=True,
        MAX_LINES=5, FOOTER_K=8, CORTAR_CLAUSULAS=True
    )

def match_titles_against_candidates(titles: list, candidates: list, *, MIN_LEN, JACCARD, REQUIRE_ANCHOR, REQUIRE_FIRST_TOKEN):
    results = []; used_candidates = set()
    titles_map = {}
    for idx, title in enumerate(titles):
        t_canon = normalized_plain_canonical(title)
        if t_canon: titles_map[t_canon] = idx
    found_map = {}
    # Exactos
    for c_idx, c in enumerate(candidates):
        cc = c['text_canon']
        if cc in titles_map and c_idx not in used_candidates:
            t_idx = titles_map[cc]
            if t_idx not in found_map:
                found_map[t_idx] = {'found': True, 'pos': c['pos'], 'obs': "Exacta", 'used_idx': c_idx}
                used_candidates.add(c_idx)
    # Robustos
    for t_idx, title in enumerate(titles):
        if t_idx in found_map: continue
        t_can = normalized_plain_canonical(title)
        if not t_can or len(t_can) < MIN_LEN: continue
        t_tokens = set(t_can.split())
        fst_t = (t_can.split()[0] if t_can else "")
        best = None
        for c_idx, c in enumerate(candidates):
            if c_idx in used_candidates: continue
            cc = c['text_canon']
            if len(cc) < MIN_LEN: continue
            if REQUIRE_ANCHOR and not has_title_anchor(cc): continue
            if c['orig'].strip().endswith('.') and len(c['orig'].split()) >= 12: continue
            if is_body_line(c['orig']): continue
            fst_c = (cc.split()[0] if cc else "")
            if REQUIRE_FIRST_TOKEN and (not fst_t or fst_t != fst_c): continue
            jac = jaccard(t_tokens, set(cc.split()))
            if jac >= JACCARD:
                score = jac + (0.02 if has_title_anchor(cc) else 0.0)
                if (best is None) or (score > best[0]) or (score == best[0] and c['pos'] < best[2]['pos']):
                    best = (score, c_idx, c)
        if best is not None:
            _, c_idx, c = best
            found_map[t_idx] = {'found': True, 'pos': c['pos'], 'obs': "Robusta", 'used_idx': c_idx}
            used_candidates.add(c_idx)
    for idx, _ in enumerate(titles):
        if idx in found_map:
            r = found_map[idx]; results.append((True, r['pos'], r['obs']))
        else:
            results.append((False, float('inf'), "No"))
    return results

def fallback_footer_titles(titles, repaired_text, already_matched, *, FOOTER_K, JACCARD, REQUIRE_ANCHOR):
    found = {}; pages = repaired_text.split(PAGE_BREAK); pos_global = 0
    for page in pages:
        rep_lines = page.split('\n')
        line_starts, acc = [], 0
        for s in rep_lines:
            line_starts.append(acc); acc += len(s) + 1
        non_empty = [i for i, ln in enumerate(rep_lines) if ln.strip()]
        footer_idx = non_empty[-FOOTER_K:] if non_empty else []
        for i in footer_idx:
            line = rep_lines[i].strip()
            if not line or is_body_line(line): continue
            line_can = normalized_plain_canonical(line); line_set = set(line_can.split())
            anchor = has_title_anchor(line_can)
            for t_idx, title in enumerate(titles):
                if already_matched[t_idx] or t_idx in found: continue
                t_can = normalized_plain_canonical(title)
                if not t_can: continue
                t_set = set(t_can.split())
                jac = jaccard(t_set, line_set)
                if (t_can in line_can) or ((REQUIRE_ANCHOR or anchor) and jac >= max(JACCARD-0.10, 0.70)) or (jac >= max(JACCARD+0.02, JACCARD)):
                    found[t_idx] = (pos_global + line_starts[i], "Footer")
        pos_global += sum(len(s) + 1 for s in rep_lines) + 1
    return found

# =========================
# Núcleo
# =========================
def extraer_clausulas_por_titulo_mejorado(pdf_path, excel_base, progress_bar, root):
    if not os.path.exists(excel_base):
        messagebox.showerror("Error", f"No se encuentra el archivo:\n{excel_base}")
        return [], 0, 0
    try:
        df = pd.read_excel(excel_base); df = normalize_headers(df)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el Excel:\n{e}")
        return [], 0, 0

    params = get_params(excel_base)

    # Validaciones mínimas (solo usamos las canónicas)
    for col in ["Cláusula", "Texto de la cláusula"]:
        if col not in df.columns:
            messagebox.showerror("Error", f"El Excel debe tener la columna '{col}'.")
            return [], 0, 0

    # Filtra filas sin título
    titles_series = df["Cláusula"].fillna("").astype(str)
    mask_nonempty = titles_series.str.strip().ne("")
    titles = titles_series[mask_nonempty].tolist()
    texts  = df.loc[mask_nonempty, "Texto de la cláusula"].fillna("").astype(str).tolist()
    obs_series = df.loc[mask_nonempty, "Observaciones."].fillna("").astype(str) if "Observaciones." in df.columns else pd.Series([""]*mask_nonempty.sum())

    try:
        repaired_text, _, _ = normalize_pdf_text(pdf_path)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el PDF:\n{e}")
        return [], 0, 0

    repaired_text, norm_lines, _ = crop_from_cp_and_clausulas(repaired_text, cortar_clausulas=params["CORTAR_CLAUSULAS"])

    candidates = build_title_candidates(norm_lines, repaired_text, max_lines_combo=params["MAX_LINES"])
    matches = match_titles_against_candidates(
        titles, candidates,
        MIN_LEN=params["MIN_LEN"], JACCARD=params["JACCARD"],
        REQUIRE_ANCHOR=params["REQUIRE_ANCHOR"], REQUIRE_FIRST_TOKEN=params["REQUIRE_FIRST_TOKEN"]
    )

    already = [f for (f, _, _) in matches]
    footer_hits = fallback_footer_titles(
        titles, repaired_text, already,
        FOOTER_K=params["FOOTER_K"], JACCARD=params["JACCARD"], REQUIRE_ANCHOR=params["REQUIRE_ANCHOR"]
    )

    final_matches = []
    for idx, (f, p, o) in enumerate(matches):
        if not f and idx in footer_hits:
            pos, obs = footer_hits[idx]; final_matches.append((True, pos, obs))
        else:
            final_matches.append((f, p, o))

    # Construcción de resultados con encabezados CANÓNICOS y orden requerido
    resultados = []
    found_counter = 1
    for idx, ((found, pos, _), title, txt, obs_base) in enumerate(zip(final_matches, titles, texts, obs_series.tolist())):
        no_val = found_counter if found else ""  # "No." solo para encontradas
        if found: found_counter += 1
        resultados.append({
            "No.": no_val,
            "Cláusula": title,
            "Texto de la cláusula": txt,
            "Tipo de operación": "",            # se deja para diligenciar
            "Valor asegurado": "",             # se deja para diligenciar
            "Observaciones.": obs_base or ""   # respeta el punto final
        })

    # Orden: encontradas primero (tienen No.), luego el resto manteniendo orden original
    resultados.sort(key=lambda r: (r["No."] == "", r["No."] if r["No."] != "" else 10**9))

    return resultados, len(titles), sum(1 for r in resultados if r["No."] != "")

# =========================
# Guardado en Excel (tabla con encabezados exactos)
# =========================
def guardar_resultados_en_excel(resultados, nombre_salida):
    df_salida = pd.DataFrame(resultados)
    # Garantiza las columnas en el orden canónico
    df_salida = df_salida[CANON_HEADERS]
    df_salida.to_excel(nombre_salida, index=False)

    wb = load_workbook(nombre_salida)
    ws = wb.active

    # Tabla
    table_ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    tab = Table(displayName="Resultados", ref=table_ref)
    style = TableStyleInfo(
        name="TableStyleLight1",
        showFirstColumn=False, showLastColumn=False,
        showRowStripes=False, showColumnStripes=False
    )
    tab.tableStyleInfo = style
    ws.add_table(tab)

    # Estilos
    font_header = Font(bold=True, color="FFFFFF", name="Calibri", size=9)
    fill_header = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
    font_body = Font(name="Calibri", size=9)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Encabezados
    for cell in ws[1]:
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center

    # Anchos
    widths = [8, 40, 100, 22, 22, 28]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Body
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for j, cell in enumerate(row, start=1):
            cell.font = font_body
            cell.alignment = align_center if j in (1,4,5) else align_left
        ws.row_dimensions[row[0].row].height = 60

    wb.save(nombre_salida)
    return nombre_salida

# =========================
# Hilo y GUI
# =========================
def run_analysis_thread(ruta_pdf, excel_base, progress_bar, root):
    try:
        root.after(0, lambda: progress_bar.config(mode="indeterminate", style="blue.Horizontal.TProgressbar"))
        root.after(0, lambda: progress_bar.start())

        resultados, total_clausulas, encontradas_count = extraer_clausulas_por_titulo_mejorado(
            ruta_pdf, excel_base, progress_bar, root
        )

        root.after(0, lambda: progress_bar.stop())
        root.after(0, lambda: progress_bar.config(mode="determinate", value=100, style="blue.Horizontal.TProgressbar"))

        if not resultados:
            root.after(100, lambda: messagebox.showwarning("Aviso", "No se generaron resultados."))
            return

        ruta_salida = elegir_ruta_guardado(ruta_pdf)
        if not ruta_salida:
            root.after(100, lambda: messagebox.showinfo("Cancelado", "Guardado cancelado por el usuario."))
            return

        try:
            guardar_resultados_en_excel(resultados, ruta_salida)
        except PermissionError:
            msg = ("No se pudo guardar el archivo.\n\n"
                   "Es posible que el archivo de destino esté ABIERTO o protegido.\n"
                   "Ciérralo o elige otra ruta y vuelve a intentarlo.")
            root.after(100, lambda m=msg: messagebox.showerror("Error de permisos", m))
            return
        except Exception as err:
            msg = f"Ocurrió un error al guardar:\n{err}"
            root.after(100, lambda m=msg: messagebox.showerror("Error", m))
            return

        no_encontradas_count = total_clausulas - encontradas_count
        porcentaje = (encontradas_count / total_clausulas) * 100 if total_clausulas > 0 else 0

        mensaje = (
            "¡Análisis completado!\n"
            f"Se guardó en:\n{ruta_salida}\n\n"
            "Estadísticas:\n"
            f"Total de cláusulas: {total_clausulas}\n"
            f"Encontradas: {encontradas_count} ({porcentaje:.1f}%)\n"
            f"No encontradas: {no_encontradas_count} ({100-porcentaje:.1f}%)"
        )
        root.after(100, lambda m=mensaje: messagebox.showinfo("¡Listo!", m))

    except Exception as err:
        msg = f"Ocurrió un error:\n{err}"
        root.after(100, lambda m=msg: messagebox.showerror("Error", m))
    finally:
        root.after(100, lambda: progress_bar.pack_forget())

def elegir_ruta_guardado(ruta_pdf: str) -> str | None:
    carpeta_pdf = os.path.dirname(ruta_pdf)
    pdf_nombre = os.path.basename(ruta_pdf)
    nombre_sin_prefijo = pdf_nombre.replace("30_Sura_", "")
    nombre_sugerido = os.path.splitext(nombre_sin_prefijo)[0] + ".xlsx"

    ruta_destino = filedialog.asksaveasfilename(
        title="Guardar resultados como...",
        initialdir=carpeta_pdf,
        initialfile=nombre_sugerido,
        defaultextension=".xlsx",
        filetypes=[("Archivo de Excel", "*.xlsx")],
        confirmoverwrite=True
    )
    return ruta_destino if ruta_destino else None

def seleccionar_pdf_y_procesar(excel_base: str, ramo_label: StringVar):
    ruta_pdf = filedialog.askopenfilename(
        title=f"Selecciona un archivo PDF ({ramo_label.get()})",
        filetypes=[("Archivos PDF", "*.pdf")]
    )
    if ruta_pdf:
        progress_bar.pack(pady=10)
        root.update_idletasks()
        t = threading.Thread(target=run_analysis_thread, args=(ruta_pdf, excel_base, progress_bar, root))
        t.daemon = True
        t.start()

def mostrar_login():
    login_win = Toplevel()
    login_win.title("Acceso")
    login_win.geometry("300x180")
    login_win.configure(bg="#0070C0")
    login_win.grab_set()

    try:
        login_win.iconbitmap("icono.ico")
    except:
        pass

    Label(login_win, text="Ingrese la contraseña", font=("Arial", 12, "bold"),
          bg="#0070C0", fg="white").pack(pady=15)

    entry = Entry(login_win, show="*", font=("Arial", 12))
    entry.pack(pady=10)
    entry.focus()

    def verificar():
        if entry.get() == "kt1324":
            login_win.destroy()
            mostrar_app()
        else:
            messagebox.showerror("Error", "Contraseña incorrecta")
            entry.delete(0, "end")

    Button(login_win, text="Ingresar", command=verificar,
           bg="white", fg="#0070C0", font=("Arial", 11, "bold"),
           padx=10, pady=5).pack(pady=10)

    login_win.bind("<Return>", lambda e: verificar())

def mostrar_app():
    global root, progress_bar
    root = Tk()
    root.title("Analizador de Cláusulas – MRC / Transporte (V 3.5)")
    try:
        root.iconbitmap('icono.ico')
    except:
        pass

    root.geometry("680x390")
    root.minsize(640, 360)

    ramo_label = StringVar(value="(sin ramo)")

    Label(root, text="Analizador de Cláusulas", font=("Arial", 15, "bold")).pack(pady=(10, 2))
    Label(root, text="Compara títulos del Excel base vs el PDF de la póliza.", font=("Arial", 10)).pack()
    Label(root, text="Versión 3.5 (encabezados canónicos + por-ramo automático)", font=("Arial", 9)).pack()

    def _go_mrc():
        ramo_label.set("MRC")
        seleccionar_pdf_y_procesar(MRC_BASE, ramo_label)

    def _go_tte():
        ramo_label.set("Transporte")
        seleccionar_pdf_y_procesar(TTE_BASE, ramo_label)

    Button(root, text="📄 Analizar MRC", command=_go_mrc, font=("Arial", 11, "bold"), bg="#0070C0", fg="white", padx=16, pady=8).pack(pady=(10, 6))
    Button(root, text="🚚 Analizar Transporte", command=_go_tte, font=("Arial", 11, "bold"), bg="#0B7F3F", fg="white", padx=16, pady=8).pack(pady=4)

    Label(root, textvariable=ramo_label, font=("Arial", 9, "italic"), fg="gray").pack(pady=(6, 2))

    style = ttk.Style(); style.theme_use('default'); style.configure("blue.Horizontal.TProgressbar", background='#0070C0', troughcolor='#e0e0e0')
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=280, mode="determinate", style="blue.Horizontal.TProgressbar")

    Label(root, text="Requisitos: base_clausulas_mrc.xlsx / base_clausulas_transporte.xlsx en la misma carpeta.", font=("Arial", 8), fg="#777", wraplength=600, justify="center").pack(pady=(8, 0))

    root.mainloop()

if __name__ == "__main__":
    temp_root = Tk(); temp_root.withdraw(); mostrar_login(); temp_root.mainloop()
