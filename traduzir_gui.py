#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tradutor de Arquivos - Checklist Colaudi
Interface gráfica para tradução de arquivos Excel e Word
mantendo formatação original (cores, layout, fontes, células mescladas)
"""

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import tkinter.font as tkfont
import os, re, json, time, threading, queue, gc, traceback, ctypes
from datetime import datetime
from pathlib import Path
from deep_translator import GoogleTranslator
import win32com.client
import pythoncom

LINGUAS = [
    "Português (BR)", "English", "Espanhol", "Francês",
    "Alemão", "Italiano", "Holandês", "Russo",
    "Chinês", "Japonês", "Coreano", "Árabe",
    "Polonês", "Turco", "Grego", "Sueco",
]

COD_LINGUAS = {
    "Português (BR)": "pt", "English": "en", "Espanhol": "es",
    "Francês": "fr", "Alemão": "de", "Italiano": "it",
    "Holandês": "nl", "Russo": "ru", "Chinês": "zh-CN",
    "Japonês": "ja", "Coreano": "ko", "Árabe": "ar",
    "Polonês": "pl", "Turco": "tr", "Grego": "el",
    "Sueco": "sv",
}

SUFIXO_IDIOMA = {
    "Português (BR)": "PT_BR", "English": "EN", "Espanhol": "ES",
    "Francês": "FR", "Alemão": "DE", "Italiano": "IT",
    "Holandês": "NL", "Russo": "RU", "Chinês": "ZH_CN",
    "Japonês": "JA", "Coreano": "KO", "Árabe": "AR",
    "Polonês": "PL", "Turco": "TR", "Grego": "EL",
    "Sueco": "SV",
}

S_NUM = re.compile(r'^[\d\s\.,;:\-\(\)\[\]\/%€$£#@!&*+=<>\\|~\'\"°ºª§@#*]+$')

ARQ_CACHE = Path(__file__).parent / "cache_traducao.json"
ARQ_CONFIG = Path(__file__).parent / "config_tradutor.json"


# ==================== CACHE ====================

class CacheManager:
    def __init__(self):
        self.dados = {}
        self.carregar()

    def carregar(self):
        if ARQ_CACHE.exists():
            try:
                self.dados.update(json.loads(ARQ_CACHE.read_text("utf-8")))
            except:
                pass

    def salvar(self):
        try:
            ARQ_CACHE.write_text(
                json.dumps(self.dados, ensure_ascii=False, indent=2), "utf-8"
            )
        except:
            pass

    def obter(self, texto):
        return self.dados.get(texto)

    def definir(self, texto, traducao):
        self.dados[texto] = traducao


# ==================== TRADUTOR ====================

class Tradutor:
    def __init__(self, lingua_destino, cache, log_callback):
        self.codigo = COD_LINGUAS.get(lingua_destino, "pt")
        self.cache = cache
        self.log = log_callback
        self._tradutor = None

    def _get_tradutor(self):
        if self._tradutor is None:
            self._tradutor = GoogleTranslator(source="auto", target=self.codigo)
        return self._tradutor

    def texto_valido(self, t):
        return (
            t and isinstance(t, str) and len(t.strip()) >= 3
            and not S_NUM.match(t.strip())
        )

    def traduzir_batch(self, textos):
        if not textos:
            return {}
        resultados = {}
        nao_cache = []
        for t in textos:
            if t in self.cache.dados:
                resultados[t] = self.cache.dados[t]
            else:
                nao_cache.append(t)
        if nao_cache:
            try:
                tradutor = self._get_tradutor()
                TAM = 100
                total = len(nao_cache)
                for i in range(0, total, TAM):
                    bloco = [t[:5000] for t in nao_cache[i:i + TAM]]
                    try:
                        traducoes = tradutor.translate_batch(bloco)
                        for orig, trad in zip(bloco, traducoes):
                            self.cache.definir(orig, trad)
                            resultados[orig] = trad
                    except Exception as e:
                        self.log(f"Erro batch {i}-{i+TAM}: {e}", "warning")
                        for t in bloco:
                            try:
                                r = tradutor.translate(t)
                                self.cache.definir(t, r)
                                resultados[t] = r
                            except:
                                resultados[t] = t
                    time.sleep(0.4)
                self.cache.salvar()
            except Exception as e:
                self.log(f"Falha na tradução: {e}", "error")
                for t in nao_cache:
                    resultados[t] = self.cache.dados.get(t, t)
        return resultados


# ==================== TRADUTORES DE ARQUIVO ====================

class TradutorArquivos:
    def __init__(self, tradutor, log_callback, progresso_callback, stop_event):
        self.tradutor = tradutor
        self.log = log_callback
        self.progresso = progresso_callback
        self.stop = stop_event

    def traduzir_excel(self, path_orig, path_dest):
        if self.stop.is_set():
            return False

        ext = os.path.splitext(path_orig)[1].lower()

        if ext == ".xlsx":
            return self._traduzir_excel_openpyxl(path_orig, path_dest)
        else:
            return self._traduzir_excel_com(path_orig, path_dest)

    def _traduzir_excel_openpyxl(self, path_orig, path_dest):
        import openpyxl
        from openpyxl.utils import get_column_letter

        wb = openpyxl.load_workbook(path_orig)
        if self.stop.is_set():
            wb.close()
            return False
        textos = set()

        for sheet_name in wb.sheetnames:
            if self.stop.is_set():
                wb.close()
                return False
            ws = wb[sheet_name]
            if self.tradutor.texto_valido(sheet_name):
                textos.add(sheet_name)
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value and self.tradutor.texto_valido(str(cell.value)):
                        textos.add(str(cell.value).strip())

        if self.stop.is_set():
            wb.close()
            return False

        traducoes = self.tradutor.traduzir_batch(list(textos))

        if self.stop.is_set():
            wb.close()
            return False

        for sheet_name in wb.sheetnames:
            if self.stop.is_set():
                wb.close()
                return False
            ws = wb[sheet_name]
            if sheet_name in traducoes:
                novo = traducoes[sheet_name][:31]
                if novo != sheet_name:
                    ws.title = novo
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value and str(cell.value).strip() in traducoes:
                        cell.value = traducoes[str(cell.value).strip()]

        wb.save(path_dest)
        wb.close()
        return True

    def _traduzir_excel_com(self, path_orig, path_dest):
        excel = None
        wb = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False

            wb = excel.Workbooks.Open(os.path.abspath(path_orig))

            textos = set()
            for sheet in wb.Worksheets:
                nome = sheet.Name
                if self.tradutor.texto_valido(nome):
                    textos.add(nome)
                used = sheet.UsedRange
                if used is None:
                    continue
                try:
                    dados = used.Value
                except:
                    continue
                if dados is None:
                    continue
                if not isinstance(dados, (list, tuple)):
                    dados = [[dados]]
                for row in dados:
                    for cell_val in row:
                        if self.tradutor.texto_valido(cell_val):
                            textos.add(str(cell_val).strip())

            if self.stop.is_set():
                wb.Close(False)
                excel.Quit()
                return False

            traducoes = self.tradutor.traduzir_batch(list(textos))

            for sheet in wb.Worksheets:
                nome = sheet.Name
                if nome in traducoes:
                    novo = traducoes[nome][:31]
                    if novo != nome:
                        try:
                            sheet.Name = novo
                        except:
                            pass

                used = sheet.UsedRange
                if used is None:
                    continue
                try:
                    r, c = used.Rows.Count, used.Columns.Count
                except:
                    continue

                for row_idx in range(1, r + 1):
                    for col_idx in range(1, c + 1):
                        if self.stop.is_set():
                            wb.Close(False)
                            excel.Quit()
                            return False
                        try:
                            cell = used.Cells(row_idx, col_idx)
                            val = cell.Value
                            if self.tradutor.texto_valido(val) and str(val).strip() in traducoes:
                                trad = traducoes[str(val).strip()]
                                if trad != val:
                                    cell.Value = trad
                        except:
                            pass

            if os.path.exists(path_dest):
                os.remove(path_dest)
            wb.SaveAs(os.path.abspath(path_dest), FileFormat=51)
            wb.Close(False)
            excel.Quit()
            return True

        except BaseException as e:
            self.log(f"  ERRO COM Excel: {e}", "error")
            if wb:
                try:
                    wb.Close(False)
                except:
                    pass
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
            return False

    def traduzir_word(self, path_orig, path_dest):
        if self.stop.is_set():
            return False
        word = None
        doc = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False

            doc = word.Documents.Open(os.path.abspath(path_orig))

            textos = set()
            for para in doc.Paragraphs:
                t = para.Range.Text
                if self.tradutor.texto_valido(t):
                    textos.add(t.strip())
            for table in doc.Tables:
                for row_idx in range(1, table.Rows.Count + 1):
                    for col_idx in range(1, table.Columns.Count + 1):
                        try:
                            t = table.Cell(row_idx, col_idx).Range.Text
                            if self.tradutor.texto_valido(t):
                                textos.add(t.strip())
                        except:
                            pass

            if self.stop.is_set():
                doc.Close(False)
                word.Quit()
                return False

            traducoes = self.tradutor.traduzir_batch(list(textos))

            for para in doc.Paragraphs:
                t = para.Range.Text
                if t.strip() in traducoes:
                    para.Range.Text = traducoes[t.strip()]

            for table in doc.Tables:
                for row_idx in range(1, table.Rows.Count + 1):
                    for col_idx in range(1, table.Columns.Count + 1):
                        if self.stop.is_set():
                            doc.Close(False)
                            word.Quit()
                            return False
                        try:
                            cell = table.Cell(row_idx, col_idx)
                            t = cell.Range.Text
                            if t.strip() in traducoes:
                                cell.Range.Text = traducoes[t.strip()]
                        except:
                            pass

            if os.path.exists(path_dest):
                os.remove(path_dest)
            doc.SaveAs(os.path.abspath(path_dest), FileFormat=0)
            doc.Close(False)
            word.Quit()
            return True

        except Exception as e:
            self.log(f"  ERRO Word: {e}", "error")
            if doc:
                try:
                    doc.Close(False)
                except:
                    pass
            if word:
                try:
                    word.Quit()
                except:
                    pass
            return False




# ==================== WORKER ====================

class WorkerThread(threading.Thread):
    def __init__(self, arquivos, lingua_destino, cache, queue_msg, stop_event,
                 traduzir_nomes=False):
        super().__init__(daemon=True)
        self.arquivos_raw = arquivos
        self.lingua_destino = lingua_destino
        self.traduzir_nomes = traduzir_nomes
        self.cache = cache
        self.queue = queue_msg
        self.stop = stop_event

    def run(self):
        pythoncom.CoInitialize()
        try:
            cod = COD_LINGUAS.get(self.lingua_destino, "pt")
            sufixo_padrao = f"_{SUFIXO_IDIOMA.get(self.lingua_destino, 'PT_BR')}"

            traducao_nomes = {}
            if self.traduzir_nomes:
                try:
                    trad_nomes = GoogleTranslator(source="auto", target=cod)
                    bases = list(dict.fromkeys(
                        os.path.splitext(nome)[0] for _, _, _, nome, _ in self.arquivos_raw
                    ))
                    if bases:
                        res = trad_nomes.translate_batch(bases)
                        for orig, t in zip(bases, res):
                            traducao_nomes[orig] = t
                except Exception as e:
                    self.queue.put(("log", f"Erro ao traduzir nomes: {e}", "warning"))

            arquivos = []
            for item_id, path, raiz, nome, tipo in self.arquivos_raw:
                base, ext = os.path.splitext(nome)
                if tipo == "excel" and ext.lower() == ".xls":
                    ext_dest = ".xlsx"
                else:
                    ext_dest = ext
                if self.traduzir_nomes and base in traducao_nomes:
                    base_trad = re.sub(r'[<>:"/\\|?*]', '', traducao_nomes[base]).strip()
                    if base_trad:
                        base = base_trad
                    sufixo = "_traduzido"
                else:
                    sufixo = sufixo_padrao
                path_dest = os.path.join(raiz, f"{base}{sufixo}{ext_dest}")
                arquivos.append((path, path_dest, tipo, item_id))

            tradutor = Tradutor(self.lingua_destino, self.cache, lambda m, t="info": self.queue.put(("log", m, t)))
            trad_arq = TradutorArquivos(tradutor, lambda m, t="info": self.queue.put(("log", m, t)),
                                        self._progresso, self.stop)

            total = len(arquivos)
            for idx, (path_orig, path_dest, tipo, item_id) in enumerate(arquivos, 1):
                if self.stop.is_set():
                    self._log("Tradução interrompida pelo usuário.", "warning")
                    self.queue.put(("status", "Interrompido"))
                    self.queue.put(("pronto", None))
                    return

                nome = os.path.basename(path_orig)
                self._log(f"[{idx}/{total}] {nome} ...", "info")
                self._progresso(idx - 1, total, nome)

                inicio = time.time()
                try:
                    if tipo == "excel":
                        ok = trad_arq.traduzir_excel(path_orig, path_dest)
                    elif tipo == "word":
                        ok = trad_arq.traduzir_word(path_orig, path_dest)
                    else:
                        ok = False

                    decorrido = time.time() - inicio
                    if ok:
                        self._log(f"  OK ({decorrido:.0f}s)", "success")
                    else:
                        self._log(f"  FALHA ({decorrido:.0f}s)", "error")
                except Exception as e:
                    self._log(f"  ERRO: {e}", "error")

                self._progresso(idx, total, nome)
                self.queue.put(("item_pronto", item_id))
                gc.collect()

            self.cache.salvar()
            self._log("Tradução concluída!", "success")
            self.queue.put(("status", "Concluído"))
            self.queue.put(("pronto", None))
        finally:
            pythoncom.CoUninitialize()

    def _log(self, msg, tipo="info"):
        self.queue.put(("log", msg, tipo))

    def _progresso(self, atual, total, arquivo=""):
        self.queue.put(("progresso", atual, total, arquivo))


# ==================== INTERFACE GRÁFICA ====================

class TradutorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Tradutor de arquivos office")
        self.root.geometry("960x740")
        self.root.minsize(750, 620)

        self._icone_caminho = Path(__file__).parent / "icone_app.ico"
        self._gerar_icone()
        if self._icone_caminho.exists():
            try:
                self.root.iconbitmap(default=str(self._icone_caminho))
            except:
                pass

        self.cache = CacheManager()
        self.queue_app = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None
        self.arquivos_selecionados = []
        self.arquivos_info = []
        self.arquivos_check = {}

        self._config_path = Path(__file__).parent / "config_tradutor.json"
        self._ultimo_dir = ""
        self._carregar_config()

        self._criar_menu()
        self._criar_widgets()
        self._poll_queue()
        self._ajustar_tema()

    def _gerar_icone(self):
        if self._icone_caminho.exists():
            return
        import struct
        W, H = 32, 32
        AZUL = (0x57, 0x79, 0x2B, 0xFF)
        BRANCO = (0xFF, 0xFF, 0xFF, 0xFF)
        CINZA = (0xE0, 0xE0, 0xE0, 0xFF)
        VERDE = (0x46, 0x73, 0x21, 0xFF)
        VERMELHO = (0x26, 0x47, 0xD2, 0xFF)
        AZUL_W = (0x9A, 0x57, 0x2B, 0xFF)
        AMARELO = (0x00, 0xA0, 0xFF, 0xFF)

        painter = {}
        for y in range(H):
            for x in range(W):
                d = (x - 16) ** 2 + (y - 16) ** 2
                dx, dy = abs(x - 16), abs(y - 16)

                if d > 240:
                    painter[(x, y)] = AZUL
                elif dy <= 10 and dx <= 8:
                    if y == 6:
                        painter[(x, y)] = VERMELHO
                    elif y == 8:
                        painter[(x, y)] = VERDE
                    elif y == 10:
                        painter[(x, y)] = AZUL_W
                    elif y == 12:
                        painter[(x, y)] = AMARELO
                    else:
                        painter[(x, y)] = BRANCO if dy <= 10 else CINZA
                elif 6 <= y < 14 and 8 < x < 24:
                    painter[(x, y)] = BRANCO
                else:
                    painter[(x, y)] = CINZA

        px = [painter[(x, H - 1 - y)] for y in range(H) for x in range(W)]
        bmp = b''.join(struct.pack('BBBB', *p) for p in px)
        row_size = ((W * 32 + 31) // 32) * 4
        pad = row_size - W * 4
        rows = []
        for y in range(H):
            rows.append(bmp[y * W * 4:(y + 1) * W * 4] + b'\x00' * pad)
        bmp_padded = b''.join(rows)
        and_mask = b'\x00' * ((W + 31) // 32) * 4 * H
        bmp_size = 40 + len(bmp_padded) + len(and_mask)
        off = 6 + 16

        with open(self._icone_caminho, 'wb') as f:
            f.write(struct.pack('<HHH', 0, 1, 1))
            f.write(struct.pack('<BBBBHHII', W, H, 0, 0, 1, 32, bmp_size, off))
            f.write(struct.pack('<IiiHHIIiiII', 40, W, H * 2, 1, 32, 0, len(bmp_padded), 0, 0, 0, 0))
            f.write(bmp_padded)
            f.write(and_mask)

    def _carregar_config(self):
        if self._config_path.exists():
            try:
                cfg = json.loads(self._config_path.read_text("utf-8"))
                self._ultimo_dir = cfg.get("ultimo_dir", "")
            except:
                pass

    def _salvar_config(self):
        try:
            self._config_path.write_text(
                json.dumps({"ultimo_dir": self.var_pasta.get()}, ensure_ascii=False),
                "utf-8",
            )
        except:
            pass

    def _ajustar_tema(self):
        fonte_padrao = tkfont.nametofont("TkDefaultFont")
        fonte_padrao.configure(size=10)
        tkfont.nametofont("TkTextFont").configure(size=10)
        tkfont.nametofont("TkHeadingFont").configure(size=11, weight="bold")
        tkfont.nametofont("TkFixedFont").configure(size=10)

        style = ttk.Style()
        try:
            style.theme_use("vista" if os.name == "nt" else "clam")
        except:
            pass
        style.configure("Accent.TButton", font=(fonte_padrao.actual("family"), 10, "bold"))
        style.configure("Treeview", rowheight=28, font=(fonte_padrao.actual("family"), 10))
        style.configure("Treeview.Heading", font=(fonte_padrao.actual("family"), 10, "bold"))
        style.configure("TLabelframe.Label", font=(fonte_padrao.actual("family"), 10, "bold"))
        style.configure("TButton", font=(fonte_padrao.actual("family"), 10), padding=(8, 4))
        style.configure("TCheckbutton", font=(fonte_padrao.actual("family"), 10))
        style.configure("TLabel", font=(fonte_padrao.actual("family"), 10))
        style.configure("TEntry", font=(fonte_padrao.actual("family"), 10))

    def _criar_menu(self):
        barra = tk.Menu(self.root)
        menu_ajuda = tk.Menu(barra, tearoff=0)
        menu_ajuda.add_command(label="Como usar", command=self._mostrar_ajuda)
        menu_ajuda.add_command(label="Sobre", command=self._mostrar_sobre)
        barra.add_cascade(label="Ajuda", menu=menu_ajuda)
        self.root.configure(menu=barra)

    def _mostrar_ajuda(self):
        top = tk.Toplevel(self.root)
        top.title("Como usar")
        if self._icone_caminho.exists():
            try:
                top.iconbitmap(default=str(self._icone_caminho))
            except:
                pass
        top.geometry("520x460")
        top.minsize(400, 320)
        top.transient(self.root)
        top.grab_set()

        txt = tk.Text(top, wrap="word", font=("Segoe UI", 10),
                      relief="sunken", borderwidth=2, padx=12, pady=12)
        txt.pack(fill="both", expand=True, padx=8, pady=8)

        scroll = ttk.Scrollbar(top, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)

        txt.insert("end", "COMO USAR O TRADUTOR DE ARQUIVOS OFFICE\n\n", "titulo")
        txt.insert("end", "1. Selecione a pasta raiz com os arquivos clicando em 📁.\n")
        txt.insert("end", "2. Escolha os tipos de arquivo desejados (Excel, Word).\n")
        txt.insert("end", "3. Selecione o idioma de destino no menu suspenso.\n")
        txt.insert('end', '4. Marque "Trad. nomes" se quiser traduzir também os nomes dos arquivos.\n')
        txt.insert("end", "5. Clique em 🔍 Escanear para listar os arquivos encontrados.\n")
        txt.insert("end", "6. Clique na coluna de seleção (☐/☑) para marcar/desmarcar arquivos.\n")
        txt.insert("end", "   Use a tecla Espaço para alternar a seleção do item em foco.\n")
        txt.insert("end", "7. Clique em ▶ Iniciar Tradução para começar.\n\n")
        txt.insert("end", "Os arquivos traduzidos serão salvos com o sufixo do idioma\n")
        txt.insert("end", "(ex.: _PT_BR) ao lado do arquivo original.\n\n")
        txt.insert("end", "Use ⏹ Parar para interromper a qualquer momento.\n")
        txt.insert("end", "O botão Limpar Cache apaga o cache de traduções\n")
        txt.insert("end", "para forçar uma retradução completa.\n\n")
        txt.insert("end", "DICA: Arquivos com _traduzido no nome são ignorados\n")
        txt.insert("end", "automaticamente para evitar retradução.", "dica")

        txt.tag_configure("titulo", font=("Segoe UI", 12, "bold"), spacing3=8)
        txt.tag_configure("dica", foreground="#0044aa", font=("Segoe UI", 10, "italic"))
        txt.configure(state="disabled")

        ttk.Button(top, text="Fechar", command=top.destroy).pack(pady=(0, 8))

    def _mostrar_sobre(self):
        msg = (
            "Tradutor de Arquivos Office v1.0\n\n"
            "Traduz conteúdo de arquivos Excel e Word\n"
            "mantendo a formatação original (cores, layout, fontes).\n\n"
            "Tecnologias: Python, Tkinter, deep-translator\n"
            "Padrões: WCAG AA, ISO 9241-115/161, HIG\n\n"
            "Desenvolvido por Autolinx Automação\n"
            "http://autolinx.me\n\n"
            "© 2026"
        )
        messagebox.showinfo("Sobre", msg)

    def _criar_widgets(self):
        # ===== VARIÁVEIS =====
        self.var_pasta = tk.StringVar(value=self._ultimo_dir)
        self.var_excel = tk.BooleanVar(value=True)
        self.var_word = tk.BooleanVar(value=True)
        self.var_lingua = tk.StringVar(value="Português (BR)")
        self.var_traduzir_nomes = tk.BooleanVar(value=False)

        # ===== FRAME TOPO =====
        frame_topo = ttk.LabelFrame(self.root, text="Configuração", padding=12)
        frame_topo.pack(fill="x", padx=12, pady=(12, 6))

        # Pasta
        row_pasta = ttk.Frame(frame_topo)
        row_pasta.pack(fill="x", pady=4)
        lbl_pasta = ttk.Label(row_pasta, text="Pasta raiz:", width=14)
        lbl_pasta.pack(side="left")
        self.entry_pasta = ttk.Entry(row_pasta, textvariable=self.var_pasta)
        self.entry_pasta.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.btn_pasta = ttk.Button(row_pasta, text="  📁  ", command=self._selecionar_pasta, width=4)
        self.btn_pasta.pack(side="left")

        # Tipos
        row_tipos = ttk.Frame(frame_topo)
        row_tipos.pack(fill="x", pady=4)
        ttk.Label(row_tipos, text="Tipos de arquivo:", width=14).pack(side="left")
        self.chk_excel = ttk.Checkbutton(row_tipos, text="Excel (.xls/.xlsx)", variable=self.var_excel)
        self.chk_excel.pack(side="left", padx=(0, 12))
        self.chk_word = ttk.Checkbutton(row_tipos, text="Word (.doc/.docx)", variable=self.var_word)
        self.chk_word.pack(side="left", padx=(0, 12))

        # Idioma + Botoes
        row_idioma = ttk.Frame(frame_topo)
        row_idioma.pack(fill="x", pady=4)
        ttk.Label(row_idioma, text="Idioma destino:", width=14).pack(side="left")
        cmb_idioma = ttk.Combobox(
            row_idioma, textvariable=self.var_lingua,
            values=LINGUAS, state="readonly", width=22,
        )
        cmb_idioma.pack(side="left", padx=(0, 12))

        self.chk_trad_nomes = ttk.Checkbutton(
            row_idioma, text="Trad. nomes", variable=self.var_traduzir_nomes,
        )
        self.chk_trad_nomes.pack(side="left", padx=(0, 12))

        self.btn_escanear = ttk.Button(row_idioma, text="  🔍 Escanear  ", command=self._escanear)
        self.btn_escanear.pack(side="left", padx=3)
        self.btn_sel_todos = ttk.Button(row_idioma, text=" Selecionar Todos ", command=self._selecionar_todos)
        self.btn_sel_todos.pack(side="left", padx=3)
        self.btn_limpar_sel = ttk.Button(row_idioma, text=" Limpar ", command=self._limpar_selecao)
        self.btn_limpar_sel.pack(side="left", padx=3)

        # ===== FRAME LISTA =====
        frame_lista = ttk.LabelFrame(self.root, text="Arquivos encontrados", padding=6)
        frame_lista.pack(fill="both", expand=True, padx=12, pady=6)

        cols = ("check", "arquivo", "pasta", "tamanho")
        self.tree = ttk.Treeview(
            frame_lista, columns=cols, show="headings",
            selectmode="none", height=6,
        )
        self.tree.heading("check", text="")
        self.tree.heading("arquivo", text="Arquivo")
        self.tree.heading("pasta", text="Pasta")
        self.tree.heading("tamanho", text="Tamanho")
        self.tree.column("check", width=36, anchor="center", stretch=False)
        self.tree.column("arquivo", width=300, anchor="w")
        self.tree.column("pasta", width=300, anchor="w")
        self.tree.column("tamanho", width=90, anchor="e", stretch=False)
        self.tree.bind("<ButtonRelease-1>", self._toggle_check)
        self.tree.bind("<space>", self._toggle_check_key)
        self.tree.bind("<FocusIn>", lambda e: self.tree.focus_set())
        self.tree.tag_configure("traduzido", background="#C5E1F5")

        scroll_y = ttk.Scrollbar(frame_lista, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        # ===== FRAME PROGRESSO =====
        frame_prog = ttk.LabelFrame(self.root, text="Progresso", padding=8)
        frame_prog.pack(fill="x", padx=12, pady=6)

        row_prog = ttk.Frame(frame_prog)
        row_prog.pack(fill="x")
        self.progresso_bar = ttk.Progressbar(row_prog, mode="determinate")
        self.progresso_bar.pack(side="left", fill="x", expand=True)
        self.lbl_progresso = ttk.Label(row_prog, text="0 / 0", width=10, anchor="center")
        self.lbl_progresso.pack(side="left", padx=(12, 0))
        self.lbl_arquivo_atual = ttk.Label(row_prog, text="", width=36, anchor="e")
        self.lbl_arquivo_atual.pack(side="left", padx=(12, 0))

        # ===== FRAME LOG =====
        frame_log = ttk.LabelFrame(self.root, text="Log", padding=6)
        frame_log.pack(fill="x", padx=12, pady=6)

        self.txt_log = tk.Text(frame_log, height=5, wrap="word", state="disabled",
                                font=("Consolas", 10), relief="sunken", borderwidth=2)
        scroll_log = ttk.Scrollbar(frame_log, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scroll_log.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        scroll_log.pack(side="right", fill="y")

        # Tags de cor (WCAG AA / ISO 9241-161)
        self.txt_log.tag_configure("info", foreground="#1a1a1a")
        self.txt_log.tag_configure("success", foreground="#0d6b0d")
        self.txt_log.tag_configure("warning", foreground="#7a5c00")
        self.txt_log.tag_configure("error", foreground="#b30000")
        self.txt_log.tag_configure("highlight", foreground="#0044aa")

        # ===== FRAME BOTÕES =====
        frame_botoes = ttk.Frame(self.root, padding=6)
        frame_botoes.pack(fill="x", padx=12, pady=(0, 12))

        self.btn_iniciar = ttk.Button(
            frame_botoes, text="  ▶ Iniciar Tradução  ",
            command=self._iniciar, style="Accent.TButton",
        )
        self.btn_iniciar.pack(side="left", padx=3)

        self.btn_parar = ttk.Button(
            frame_botoes, text="  ⏹ Parar  ",
            command=self._parar, state="disabled",
        )
        self.btn_parar.pack(side="left", padx=3)

        ttk.Button(
            frame_botoes, text=" Limpar Cache ",
            command=self._limpar_cache,
        ).pack(side="right", padx=3)

        self.lbl_status = ttk.Label(frame_botoes, text="Pronto", foreground="#444", font=(None, 10, "bold"))
        self.lbl_status.pack(side="right", padx=16)

    def _selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecionar pasta raiz")
        if pasta:
            self.var_pasta.set(pasta)
            self._salvar_config()

    def _escanear(self):
        pasta = self.var_pasta.get().strip()
        if not pasta or not os.path.isdir(pasta):
            messagebox.showwarning("Aviso", "Selecione uma pasta válida.")
            return

        self._log("Escaneando arquivos...", "info")
        self.arquivos_info.clear()
        self.arquivos_check.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)

        incluir_excel = self.var_excel.get()
        incluir_word = self.var_word.get()

        encontrados = 0
        for raiz, _, arquivos in os.walk(pasta):
            for f in sorted(arquivos):
                fl = f.lower()
                if "_traduzido" in fl:
                    continue
                path = os.path.join(raiz, f)

                tipo = None
                if incluir_excel and (fl.endswith(".xls") or fl.endswith(".xlsx")):
                    tipo = "excel"
                elif incluir_word and (fl.endswith(".doc") or fl.endswith(".docx")):
                    tipo = "word"

                if tipo is None:
                    continue

                nome_pasta = os.path.relpath(raiz, pasta)
                if nome_pasta == ".":
                    nome_pasta = ""
                tamanho = os.path.getsize(path)
                if tamanho < 1024:
                    str_tam = f"{tamanho} B"
                elif tamanho < 1024 * 1024:
                    str_tam = f"{tamanho // 1024} KB"
                else:
                    str_tam = f"{tamanho / (1024*1024):.1f} MB"

                item_id = self.tree.insert("", "end", values=("☐", f, nome_pasta, str_tam))
                self.arquivos_check[item_id] = True
                self.arquivos_info.append((item_id, path, raiz, f, tipo))
                encontrados += 1

        self._selecionar_todos()
        self._log(f"{encontrados} arquivos encontrados.", "success" if encontrados else "warning")
        if not encontrados:
            self._log("Nenhum arquivo corresponde aos filtros selecionados.", "warning")

    def _toggle_check(self, event):
        item = self.tree.identify_row(event.y)
        if not item or item not in self.arquivos_check:
            return
        col = self.tree.identify_column(event.x)
        if col != "#1":
            return
        self.arquivos_check[item] = not self.arquivos_check[item]
        self.tree.set(item, "check", "☑" if self.arquivos_check[item] else "☐")

    def _toggle_check_key(self, event):
        item = self.tree.focus()
        if item and item in self.arquivos_check:
            self.arquivos_check[item] = not self.arquivos_check[item]
            self.tree.set(item, "check", "☑" if self.arquivos_check[item] else "☐")
            return "break"

    def _selecionar_todos(self):
        for item in self.arquivos_check:
            self.arquivos_check[item] = True
            self.tree.set(item, "check", "☑")

    def _limpar_selecao(self):
        for item in self.arquivos_check:
            self.arquivos_check[item] = False
            self.tree.set(item, "check", "☐")

    def _iniciar(self):
        if not self.arquivos_info:
            messagebox.showwarning("Aviso", "Escaneie os arquivos primeiro.")
            return

        selecionados = []
        for item_id, path, raiz, nome, tipo in self.arquivos_info:
            if self.arquivos_check.get(item_id, False):
                selecionados.append((item_id, path, raiz, nome, tipo))

        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos um arquivo.")
            return

        self.arquivos_selecionados = selecionados
        self.stop_event.clear()

        # Reset UI
        self.progresso_bar["value"] = 0
        self.lbl_progresso["text"] = f"0 / {len(selecionados)}"
        self.lbl_arquivo_atual["text"] = ""
        self.btn_iniciar["state"] = "disabled"
        self.btn_parar["state"] = "normal"
        self.lbl_status["text"] = "Traduzindo..."
        self.lbl_status["foreground"] = "#0077C8"

        lingua = self.var_lingua.get()
        self.worker = WorkerThread(
            selecionados, lingua, self.cache, self.queue_app, self.stop_event,
            traduzir_nomes=self.var_traduzir_nomes.get(),
        )
        self.worker.start()

    def _parar(self):
        if self.worker and self.worker.is_alive():
            self.stop_event.set()
            ctypes.pythonapi.PyThreadState_SetAsyncExc(
                ctypes.c_long(self.worker.ident),
                ctypes.py_object(SystemExit),
            )
            self._log("Parada solicitada.", "warning")
        self.btn_parar["state"] = "disabled"
        self.btn_iniciar["state"] = "normal"
        self.lbl_status["text"] = "Interrompido"
        self.lbl_status["foreground"] = "#b8860b"
        self.progresso_bar["value"] = 0

    def _poll_queue(self):
        try:
            while True:
                msg = self.queue_app.get_nowait()
                tipo = msg[0]

                if tipo == "log":
                    _, texto, subtipo = msg
                    self._log(texto, subtipo)
                elif tipo == "progresso":
                    _, atual, total, arquivo = msg
                    self.progresso_bar["value"] = (atual / total) * 100 if total else 0
                    self.lbl_progresso["text"] = f"{atual} / {total}"
                    self.lbl_arquivo_atual["text"] = os.path.basename(arquivo) if arquivo else ""
                elif tipo == "status":
                    _, texto = msg
                    self.lbl_status["text"] = texto
                    if texto == "Concluído":
                        self.lbl_status["foreground"] = "#1a8a1a"
                    elif texto == "Interrompido":
                        self.lbl_status["foreground"] = "#b8860b"
                elif tipo == "item_pronto":
                    _, item_id = msg
                    self.tree.item(item_id, tags=("traduzido",))
                elif tipo == "pronto":
                    self.btn_iniciar["state"] = "normal"
                    self.btn_parar["state"] = "disabled"
                    if not self.stop_event.is_set():
                        messagebox.showinfo("Concluído", "Tradução finalizada com sucesso!")
        except queue.Empty:
            pass

        self.root.after(100, self._poll_queue)

    def _log(self, texto, tipo="info"):
        self.txt_log.configure(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{timestamp}] {texto}\n", tipo)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _limpar_cache(self):
        if messagebox.askyesno("Limpar Cache", "Limpar cache de traduções?"):
            self.cache.dados.clear()
            self.cache.salvar()
            self._log(f"Cache limpo.", "warning")

    def executar(self):
        self.root.mainloop()


# ==================== MAIN ====================

if __name__ == "__main__":
    app = TradutorApp()
    app.executar()
