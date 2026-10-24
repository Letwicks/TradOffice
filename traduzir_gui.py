#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tradutor de Arquivos Office
Interface gráfica para tradução em lote de arquivos Excel e Word
mantendo a formatação original (cores, layout, fontes, células mescladas).
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont
import os, re, json, time, threading, queue, gc, traceback, sys
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse
import urllib.request
import webbrowser

from deep_translator import GoogleTranslator

try:
    import win32com.client
    import pythoncom
    TEM_COM = True
except ImportError:  # fora do Windows / sem pywin32: só .xlsx via openpyxl
    win32com = pythoncom = None
    TEM_COM = False

VERSAO = "1.2"
REPO_API = "https://api.github.com/repos/Letwicks/TradOffice/releases/latest"
REPO_URL = "https://github.com/Letwicks/TradOffice/releases"

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

ORIGEM_AUTO = "Detectar automaticamente"
LINGUAS_ORIGEM = [ORIGEM_AUTO] + LINGUAS

EXT_EXCEL = (".xls", ".xlsx")
EXT_WORD = (".doc", ".docx")

# Sufixo que o próprio programa acrescenta aos arquivos gerados (ex.: _PT_BR)
RE_SUFIXO_SAIDA = re.compile(
    r"_(%s)$" % "|".join(sorted(SUFIXO_IDIOMA.values(), key=len, reverse=True)),
    re.IGNORECASE,
)
RE_URL = re.compile(r"^(https?://|www\.)\S+$|^[\w.+-]+@[\w-]+(\.[\w-]+)+$", re.IGNORECASE)
# Códigos sem minúsculas e com ao menos um dígito: M55.780013, A1, PLC-01, 12B
RE_CODIGO = re.compile(r"^(?=[^a-z]*\d)[A-Z0-9]+([._\-/][A-Z0-9]+)*$")
RE_NOME_INVALIDO = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _dir_app():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _dir_dados():
    """Pasta gravável para cache/config (no .exe, __file__ aponta para uma pasta temporária)."""
    pasta = _dir_app()
    if getattr(sys, "frozen", False) and not os.access(pasta, os.W_OK):
        pasta = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "TradOffice"
        pasta.mkdir(parents=True, exist_ok=True)
    return pasta


DIR_APP = _dir_app()
DIR_DADOS = _dir_dados()
ARQ_CACHE = DIR_DADOS / "cache_traducao.json"
ARQ_CONFIG = DIR_DADOS / "config_tradutor.json"


class Interrompido(Exception):
    """Usuário pediu para parar."""


class ServicoIndisponivel(Exception):
    """O serviço de tradução falhou repetidamente (sem rede, bloqueio por excesso de requisições...)."""


# ==================== UTILITÁRIOS DE TEXTO ====================

def limpar(t):
    """Remove espaços e marcas de fim de célula do Word (\\x07) das pontas."""
    return re.sub(r"^[\s\x07]+|[\s\x07]+$", "", t)


def texto_valido(t):
    """Indica se vale a pena enviar o texto para tradução."""
    if not isinstance(t, str):
        return False
    s = limpar(t)
    if len(s) < 2 or not any(c.isalpha() for c in s):
        return False
    return not (RE_URL.match(s) or RE_CODIGO.match(s))


def com_espacos(orig, trad):
    """Reaplica os espaços/quebras das pontas do texto original na tradução."""
    inicio = re.match(r"^\s*", orig).group()
    fim = re.search(r"\s*$", orig).group()
    return inicio + trad + fim


def nome_aba_seguro(nome, usados):
    """Nome de aba válido no Excel: sem []:*?/\\, até 31 caracteres e sem repetir `usados` (minúsculas)."""
    nome = re.sub(r"[\[\]:*?/\\]", "", nome).strip().strip("'").strip()[:31]
    if not nome:
        return None
    base, n = nome, 2
    while nome.lower() in usados:
        sufixo = f" ({n})"
        nome = base[:31 - len(sufixo)] + sufixo
        n += 1
    return nome


def versao_tupla(v):
    return tuple(int(x) for x in re.findall(r"\d+", v or ""))


# ==================== ARQUIVOS ====================

def eh_arquivo_traduzido(nome, nomes_na_pasta, sufixo_destino):
    """Reconhece arquivos gerados por este programa para não traduzi-los de novo."""
    if "_traduzido" in nome.lower():
        return True
    base = os.path.splitext(nome)[0]
    m = RE_SUFIXO_SAIDA.search(base)
    if not m:
        return False
    if m.group(1).upper() == sufixo_destino.upper():
        return True
    # Outro idioma: só é saída se o original (sem o sufixo) estiver na mesma pasta
    original = base[:m.start()].lower()
    return any(os.path.splitext(n)[0].lower() == original for n in nomes_na_pasta)


def escanear_pasta(pasta, incluir_excel=True, incluir_word=True, sufixo_destino="PT_BR"):
    """Retorna [(caminho, raiz, nome, tipo)] dos arquivos a traduzir."""
    encontrados = []
    for raiz, dirs, arquivos in os.walk(pasta):
        dirs.sort(key=str.lower)
        for f in sorted(arquivos, key=str.lower):
            fl = f.lower()
            if fl.startswith("~$"):  # arquivo de bloqueio do Office
                continue
            ext = os.path.splitext(fl)[1]
            if incluir_excel and ext in EXT_EXCEL:
                tipo = "excel"
            elif incluir_word and ext in EXT_WORD:
                tipo = "word"
            else:
                continue
            if eh_arquivo_traduzido(f, arquivos, sufixo_destino):
                continue
            encontrados.append((os.path.join(raiz, f), raiz, f, tipo))
    return encontrados


def planejar_destinos(arquivos, sufixo, traducao_nomes=None):
    """
    arquivos: [(item_id, caminho, raiz, nome, tipo)]
    Retorna [(item_id, caminho, caminho_destino, tipo)] sem destinos repetidos.
    """
    usados = set()
    planos = []
    for item_id, path, raiz, nome, tipo in arquivos:
        base, ext = os.path.splitext(nome)
        ext_dest = ".xlsx" if tipo == "excel" else ext
        if traducao_nomes and base in traducao_nomes:
            base_trad = RE_NOME_INVALIDO.sub("", traducao_nomes[base]).strip().rstrip(". ")
            if base_trad:
                base = base_trad
        final = f"_{sufixo}{ext_dest}"
        limite = 240 - len(os.path.join(raiz, final))
        if len(base) > limite and limite > 10:
            base = base[:limite]
        destino = os.path.join(raiz, base + final)
        n = 2
        while destino.lower() in usados:
            destino = os.path.join(raiz, f"{base} ({n}){final}")
            n += 1
        usados.add(destino.lower())
        planos.append((item_id, path, destino, tipo))
    return planos


def formatar_tamanho(tamanho):
    if tamanho < 1024:
        return f"{tamanho} B"
    if tamanho < 1024 * 1024:
        return f"{tamanho // 1024} KB"
    return f"{tamanho / (1024 * 1024):.1f} MB"


# ==================== CACHE ====================

class CacheManager:
    """Cache de traduções separado por idioma de destino: {"pt": {texto: tradução}, ...}."""

    VERSAO = 2

    def __init__(self, arquivo=ARQ_CACHE):
        self.arquivo = Path(arquivo)
        self.dados = {}
        self.lock = threading.RLock()
        self._alterado = False
        self.carregar()

    def carregar(self):
        if not self.arquivo.exists():
            return
        try:
            bruto = json.loads(self.arquivo.read_text("utf-8"))
        except (OSError, ValueError):
            # Guarda o arquivo danificado em vez de sobrescrevê-lo com um cache vazio
            try:
                self.arquivo.replace(self.arquivo.with_name(self.arquivo.stem + ".corrompido.json"))
            except OSError:
                pass
            return
        if not isinstance(bruto, dict):
            return
        if bruto.get("__versao__") == self.VERSAO:
            idiomas = bruto.get("idiomas", {})
        else:
            # Formato antigo (v1.1): dicionário plano, sem idioma. Era gerado com destino pt.
            idiomas = {"pt": bruto}
            self._alterado = True
        with self.lock:
            for cod, tabela in idiomas.items():
                if not isinstance(tabela, dict):
                    continue
                destino = self.dados.setdefault(cod, {})
                for orig, trad in tabela.items():
                    if isinstance(orig, str) and isinstance(trad, str) and trad.strip():
                        destino[limpar(orig)] = trad

    def salvar(self):
        with self.lock:
            if not self._alterado:
                return
            conteudo = json.dumps({"__versao__": self.VERSAO, "idiomas": self.dados},
                                  ensure_ascii=False, indent=1)
            tmp = self.arquivo.with_name(self.arquivo.name + ".tmp")
            try:
                tmp.write_text(conteudo, "utf-8")
                os.replace(tmp, self.arquivo)  # escrita atômica
                self._alterado = False
            except OSError:
                pass

    def obter(self, cod, texto):
        with self.lock:
            return self.dados.get(cod, {}).get(texto)

    def definir(self, cod, texto, traducao):
        with self.lock:
            self.dados.setdefault(cod, {})[texto] = traducao
            self._alterado = True

    def limpar(self):
        with self.lock:
            self.dados.clear()
            self._alterado = True
        self.salvar()

    def total(self):
        with self.lock:
            return sum(len(t) for t in self.dados.values())


# ==================== TRADUTOR ====================

class Tradutor:
    """
    Traduz textos com cache. Para reduzir requisições, textos curtos de uma linha são
    agrupados (separados por quebra de linha) numa só chamada ao Google Tradutor; se a
    resposta não vier com o mesmo número de linhas, o grupo é traduzido item a item.
    """

    MAX_BYTES = 1800       # por requisição (o texto vai na URL)
    MAX_TENTATIVAS = 4
    ESPERA_INICIAL = 2.0   # segundos; dobra a cada nova tentativa
    MAX_FALHAS_SEGUIDAS = 3
    INTERVALO_SALVAR = 60  # segundos entre gravações do cache durante a tradução

    def __init__(self, lingua_destino, cache, log_callback=None, stop_event=None,
                 lingua_origem=ORIGEM_AUTO, agrupar=True, motor=None, pausa=0.4):
        self.codigo = COD_LINGUAS.get(lingua_destino, "pt")
        self.origem = COD_LINGUAS.get(lingua_origem, "auto")
        self.cache = cache
        self.log = log_callback or (lambda m, t="info": None)
        self.stop = stop_event or threading.Event()
        self.agrupar = agrupar
        self.pausa = pausa
        self._motor = motor
        self._ultima_req = 0.0
        self._falhas_seguidas = 0

    def _get_motor(self):
        if self._motor is None:
            self._motor = GoogleTranslator(source=self.origem, target=self.codigo)
        return self._motor

    def _checar_parada(self):
        if self.stop.is_set():
            raise Interrompido()

    def traduzir_batch(self, textos, progresso=None):
        """Retorna {texto: tradução}. Textos que não puderam ser traduzidos ficam de fora."""
        resultados = {}
        pendentes = []
        for t in dict.fromkeys(textos):
            trad = self.cache.obter(self.codigo, t)
            if trad is not None:
                resultados[t] = trad
            else:
                pendentes.append(t)
        if not pendentes:
            return resultados

        self.log(f"  {len(pendentes)} texto(s) novo(s) para traduzir, "
                 f"{len(resultados)} do cache", "info")
        total, feitos = len(pendentes), 0
        ultimo_save = time.monotonic()
        try:
            for grupo in self._agrupar(pendentes):
                for orig, trad in zip(grupo, self._traduzir_grupo(grupo)):
                    if trad:
                        self.cache.definir(self.codigo, orig, trad)
                        resultados[orig] = trad
                feitos += len(grupo)
                if progresso:
                    progresso(feitos, total)
                if time.monotonic() - ultimo_save > self.INTERVALO_SALVAR:
                    self.cache.salvar()
                    ultimo_save = time.monotonic()
        finally:
            self.cache.salvar()

        faltando = total - sum(1 for t in pendentes if t in resultados)
        if faltando:
            self.log(f"  Aviso: {faltando} texto(s) sem tradução foram mantidos no original.", "warning")
        return resultados

    def _agrupar(self, textos):
        grupo, tam = [], 0
        for t in textos:
            n = len(t.encode("utf-8")) + 1
            if not self.agrupar or "\n" in t or "\r" in t or n > self.MAX_BYTES:
                if grupo:
                    yield grupo
                    grupo, tam = [], 0
                yield [t]
                continue
            if grupo and tam + n > self.MAX_BYTES:
                yield grupo
                grupo, tam = [], 0
            grupo.append(t)
            tam += n
        if grupo:
            yield grupo

    def _traduzir_grupo(self, grupo):
        if len(grupo) == 1:
            return [self._traduzir_texto(grupo[0])]
        resposta = self._chamar("\n".join(grupo))
        linhas = [l.strip() for l in resposta.split("\n")] if resposta else []
        if len(linhas) != len(grupo):
            return [self._traduzir_texto(t) for t in grupo]
        saida = []
        for orig, trad in zip(grupo, linhas):
            # Linha vazia ou idêntica: pode ter sido ignorada por estar em outro idioma
            # que o detectado para o grupo; tenta sozinha.
            if not trad or trad == orig:
                trad = self._traduzir_texto(orig)
            saida.append(trad)
        return saida

    def _traduzir_texto(self, texto):
        if len(texto.encode("utf-8")) <= self.MAX_BYTES:
            return self._chamar(texto)
        partes = []
        for pedaco in self._dividir(texto):
            trad = self._chamar(limpar(pedaco)) if texto_valido(pedaco) else None
            partes.append(com_espacos(pedaco, trad) if trad else pedaco)
        return "".join(partes)

    def _dividir(self, texto):
        """Quebra textos longos em pedaços de até MAX_BYTES, preferindo fins de frase/linha."""
        corte = max(1, self.MAX_BYTES // 4)
        tokens = []
        for tok in re.split(r"(?<=[\n.!?;。！？])", texto):
            while len(tok.encode("utf-8")) > self.MAX_BYTES:
                tokens.append(tok[:corte])
                tok = tok[corte:]
            if tok:
                tokens.append(tok)
        pedacos, atual = [], ""
        for tok in tokens:
            if atual and len((atual + tok).encode("utf-8")) > self.MAX_BYTES:
                pedacos.append(atual)
                atual = tok
            else:
                atual += tok
        if atual:
            pedacos.append(atual)
        return pedacos

    def _chamar(self, texto):
        """Uma requisição, respeitando o intervalo mínimo e com novas tentativas."""
        espera = self.ESPERA_INICIAL
        for tentativa in range(1, self.MAX_TENTATIVAS + 1):
            self._checar_parada()
            falta = self._ultima_req + self.pausa - time.monotonic()
            if falta > 0 and self.stop.wait(falta):
                raise Interrompido()
            try:
                r = self._get_motor().translate(texto)
                self._ultima_req = time.monotonic()
                self._falhas_seguidas = 0
                return r.strip() if isinstance(r, str) and r.strip() else None
            except Exception as e:
                self._ultima_req = time.monotonic()
                nome = type(e).__name__
                if nome in ("TranslationNotFound", "NotValidPayload", "NotValidLength"):
                    return None
                if tentativa < self.MAX_TENTATIVAS:
                    self.log(f"  Falha na requisição ({nome}); nova tentativa em {espera:.0f}s", "warning")
                    if self.stop.wait(espera):
                        raise Interrompido()
                    espera *= 2
        self._falhas_seguidas += 1
        if self._falhas_seguidas >= self.MAX_FALHAS_SEGUIDAS:
            raise ServicoIndisponivel(
                "O Google Tradutor não está respondendo (sem internet ou limite de "
                "requisições atingido). Aguarde alguns minutos e tente novamente."
            )
        return None


# ==================== OFFICE (COM) ====================

class SessaoOffice:
    """
    Instâncias próprias e invisíveis do Excel/Word, reaproveitadas entre arquivos.
    DispatchEx sempre cria um processo novo: o Excel/Word que o usuário estiver
    usando não é tocado (Dispatch reutilizaria e, ao final, fecharia o dele).
    """

    def __init__(self):
        self._apps = {}

    def obter(self, progid):
        app = self._apps.get(progid)
        if app is not None:
            try:
                app.Name  # ainda vivo?
                return app
            except Exception:
                self._apps.pop(progid, None)
        app = win32com.client.DispatchEx(progid)
        app.Visible = False
        app.DisplayAlerts = False
        try:
            app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable: não roda macros
        except Exception:
            pass
        if progid.startswith("Excel"):
            for prop, valor in (("ScreenUpdating", False), ("EnableEvents", False),
                                ("AskToUpdateLinks", False)):
                try:
                    setattr(app, prop, valor)
                except Exception:
                    pass
        self._apps[progid] = app
        return app

    def fechar(self):
        for app in self._apps.values():
            try:
                app.Quit()
            except Exception:
                pass
        self._apps.clear()


def _matriz(valor):
    """Normaliza o retorno de Range.Value/Formula em lista de linhas."""
    if valor is None:
        return []
    if not isinstance(valor, (list, tuple)):
        return [[valor]]
    return [list(l) if isinstance(l, (list, tuple)) else [l] for l in valor]


def _texto_para_celula(texto):
    """Evita que o Excel interprete a tradução como fórmula ou número."""
    if texto[:1] in ("=", "+", "-", "@") or not any(c.isalpha() for c in texto):
        return "'" + texto
    return texto


def _stories_word(doc):
    """Corpo, cabeçalhos, rodapés, caixas de texto, notas e comentários."""
    for story in doc.StoryRanges:
        while story is not None:
            yield story
            try:
                story = story.NextStoryRange
            except Exception:
                story = None


def _range_tem_objetos(rng):
    """Substituir o texto deste trecho apagaria campos, imagens, notas etc."""
    for colecao in ("Fields", "InlineShapes", "ShapeRange", "Footnotes",
                    "Endnotes", "ContentControls"):
        try:
            if getattr(rng, colecao).Count:
                return True
        except Exception:
            pass
    return False


# ==================== TRADUTORES DE ARQUIVO ====================

class TradutorArquivos:
    def __init__(self, tradutor, log_callback, office, stop_event, progresso_textos=None):
        self.tradutor = tradutor
        self.log = log_callback
        self.office = office
        self.stop = stop_event
        self.progresso_textos = progresso_textos

    def _checar_parada(self):
        if self.stop.is_set():
            raise Interrompido()

    def _traduzir(self, textos):
        return self.tradutor.traduzir_batch(textos, self.progresso_textos)

    @staticmethod
    def _remover_destino(path_dest):
        if os.path.exists(path_dest):
            try:
                os.remove(path_dest)
            except PermissionError:
                raise PermissionError(f"O arquivo de destino está aberto em outro programa: {path_dest}")

    # ---------- Excel ----------

    def traduzir_excel(self, path_orig, path_dest):
        self._checar_parada()
        excel = None
        if TEM_COM:
            try:
                excel = self.office.obter("Excel.Application")
            except Exception as e:
                self.log(f"  Microsoft Excel indisponível: {e}", "warning")
        if excel is not None:
            return self._traduzir_excel_com(excel, path_orig, path_dest)
        if path_orig.lower().endswith(".xlsx"):
            self.log("  Usando openpyxl: imagens e gráficos podem não ser preservados.", "warning")
            return self._traduzir_excel_openpyxl(path_orig, path_dest)
        self.log("  ERRO: arquivos .xls exigem o Microsoft Excel instalado.", "error")
        return False

    def _traduzir_excel_com(self, excel, path_orig, path_dest):
        wb = excel.Workbooks.Open(
            Filename=os.path.abspath(path_orig), UpdateLinks=0, ReadOnly=True,
            IgnoreReadOnlyRecommended=True, AddToMru=False,
            Password="\u00a7",  # ignorado se não houver senha; evita travar no pedido de senha
        )
        calculo = None
        try:
            try:
                calculo = excel.Calculation
                excel.Calculation = -4135  # xlCalculationManual durante a edição
            except Exception:
                pass

            planilhas = list(wb.Worksheets)
            textos, celulas, formas = [], [], []
            for ws in planilhas:
                self._checar_parada()
                if texto_valido(ws.Name):
                    textos.append(limpar(ws.Name))
                used = ws.UsedRange
                # Leitura em bloco (1 chamada COM por planilha em vez de 1 por célula)
                valores = _matriz(used.Value)
                formulas = _matriz(used.Formula)
                da_planilha = []
                for i, linha in enumerate(valores):
                    for j, v in enumerate(linha):
                        if not texto_valido(v):
                            continue
                        try:
                            f = formulas[i][j]
                        except IndexError:
                            f = None
                        if isinstance(f, str) and f.startswith("="):
                            continue  # resultado de fórmula: não sobrescrever
                        da_planilha.append((j, i, v))
                # Ordem por coluna: textos da mesma coluna (geralmente mesmo idioma) ficam juntos
                da_planilha.sort(key=lambda x: (x[0], x[1]))
                for j, i, v in da_planilha:
                    celulas.append((used, i + 1, j + 1, v))
                    textos.append(limpar(v))
                for shp in self._formas_excel(ws):
                    textos.append(limpar(shp[1]))
                    formas.append(shp)

            traducoes = self._traduzir(textos)

            falhas = 0
            for n, (used, i, j, v) in enumerate(celulas):
                if n % 50 == 0:
                    self._checar_parada()
                trad = traducoes.get(limpar(v))
                if trad and trad != limpar(v):
                    try:
                        used.Cells(i, j).Value = _texto_para_celula(com_espacos(v, trad))
                    except Exception:
                        falhas += 1
            for texto_range, texto in formas:
                trad = traducoes.get(limpar(texto))
                if trad and trad != limpar(texto):
                    try:
                        texto_range.Text = com_espacos(texto, trad).replace("\n", "\r")
                    except Exception:
                        falhas += 1
            if falhas:
                self.log(f"  Aviso: {falhas} célula(s)/forma(s) não puderam ser alteradas "
                         "(planilha protegida?).", "warning")

            # Renomear abas por último; o Excel atualiza as fórmulas que as referenciam
            nomes = [ws.Name for ws in planilhas]
            for idx, ws in enumerate(planilhas):
                trad = traducoes.get(limpar(nomes[idx]))
                if not trad:
                    continue
                usados = {n.lower() for k, n in enumerate(nomes) if k != idx}
                novo = nome_aba_seguro(trad, usados)
                if novo and novo != nomes[idx]:
                    try:
                        ws.Name = novo
                        nomes[idx] = novo
                    except Exception:
                        pass

            if calculo is not None:  # restaura antes de salvar (o modo de cálculo vai no arquivo)
                excel.Calculation = calculo
                calculo = None
            self._remover_destino(path_dest)
            wb.SaveAs(Filename=os.path.abspath(path_dest), FileFormat=51)  # .xlsx
            return True
        finally:
            if calculo is not None:
                try:
                    excel.Calculation = calculo
                except Exception:
                    pass
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass

    @staticmethod
    def _formas_excel(ws):
        """[(TextRange, texto)] das caixas de texto/formas com texto da planilha."""
        resultado = []
        try:
            shapes = list(ws.Shapes)
        except Exception:
            return resultado
        for shp in shapes:
            try:
                tf = shp.TextFrame2
                if not tf.HasText:
                    continue
                rng = tf.TextRange
                texto = (rng.Text or "").replace("\r\n", "\n").replace("\r", "\n")
            except Exception:
                continue  # controles de formulário, grupos, imagens...
            if texto_valido(texto):
                resultado.append((rng, texto))
        return resultado

    def _traduzir_excel_openpyxl(self, path_orig, path_dest):
        import openpyxl

        wb = openpyxl.load_workbook(path_orig)
        try:
            textos, celulas, formulas = [], [], []
            for ws in wb.worksheets:
                self._checar_parada()
                if texto_valido(ws.title):
                    textos.append(limpar(ws.title))
                da_planilha = []
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.data_type == "f":
                            formulas.append(str(cell.value))
                        elif isinstance(cell.value, str) and texto_valido(cell.value):
                            da_planilha.append(cell)
                da_planilha.sort(key=lambda c: (c.column, c.row))
                for cell in da_planilha:
                    celulas.append(cell)
                    textos.append(limpar(cell.value))

            traducoes = self._traduzir(textos)
            self._checar_parada()

            for cell in celulas:
                trad = traducoes.get(limpar(cell.value))
                if trad and trad != limpar(cell.value):
                    cell.value = com_espacos(cell.value, trad)
                    cell.data_type = "s"  # nunca virar fórmula, mesmo se começar com "="

            # openpyxl não atualiza referências ao renomear: só renomeia abas não referenciadas
            try:
                formulas += [str(d.attr_text) for d in wb.defined_names.values()]
            except Exception:
                pass
            referencias = "\n".join(formulas).lower()
            for ws in wb.worksheets:
                trad = traducoes.get(limpar(ws.title))
                if not trad:
                    continue
                if ws.title.lower() in referencias:
                    self.log(f"  Aba '{ws.title}' não renomeada: é usada em fórmulas.", "warning")
                    continue
                usados = {w.title.lower() for w in wb.worksheets if w is not ws}
                novo = nome_aba_seguro(trad, usados)
                if novo and novo != ws.title:
                    ws.title = novo

            self._remover_destino(path_dest)
            wb.save(path_dest)
            return True
        finally:
            wb.close()

    # ---------- Word ----------

    def traduzir_word(self, path_orig, path_dest):
        self._checar_parada()
        if not TEM_COM:
            self.log("  ERRO: arquivos Word exigem o Microsoft Word instalado (Windows).", "error")
            return False
        word = self.office.obter("Word.Application")
        doc = word.Documents.Open(
            FileName=os.path.abspath(path_orig), ConfirmConversions=False, ReadOnly=True,
            AddToRecentFiles=False, PasswordDocument="\u00a7", Visible=False,
        )
        try:
            try:
                doc.TrackRevisions = False  # não registrar a tradução como alterações controladas
            except Exception:
                pass

            alvos, textos = [], []
            com_objetos = 0
            for story in _stories_word(doc):
                for para in story.Paragraphs:
                    self._checar_parada()
                    rng = para.Range
                    bruto = rng.Text or ""
                    if not texto_valido(bruto):
                        continue
                    if bruto.endswith(("\r", "\x07")):
                        # Exclui a marca de parágrafo/fim de célula: substituí-la juntaria parágrafos
                        rng = rng.Duplicate
                        rng.MoveEnd(1, -1)  # wdCharacter
                        texto = rng.Text or ""
                    else:
                        texto = bruto
                    partes = re.split(r"([\x0b\x0c])", texto)  # quebras de linha/página manuais
                    validas = [limpar(p) for p in partes if texto_valido(p)]
                    if not validas:
                        continue
                    if _range_tem_objetos(rng):
                        com_objetos += 1  # substituir o texto apagaria esses elementos
                        continue
                    alvos.append((rng, texto, partes))
                    textos.extend(validas)

            traducoes = self._traduzir(textos)

            def traduzir_parte(p):
                trad = traducoes.get(limpar(p)) if texto_valido(p) else None
                if not trad:
                    return p
                return com_espacos(p, re.sub(r"\r\n|\r|\n", "\x0b", trad))

            falhas = 0
            # Do fim para o início: as alterações não deslocam os trechos ainda pendentes
            for rng, texto, partes in reversed(alvos):
                self._checar_parada()
                novo = "".join(traduzir_parte(p) for p in partes)
                if novo == texto:
                    continue
                try:
                    rng.Text = novo
                except Exception:
                    falhas += 1
            if com_objetos:
                self.log(f"  Aviso: {com_objetos} parágrafo(s) com imagens/campos/notas mantidos "
                         "no original para não perder esses elementos.", "warning")
            if falhas:
                self.log(f"  Aviso: {falhas} parágrafo(s) não puderam ser alterados.", "warning")

            self._remover_destino(path_dest)
            formato = 16 if path_dest.lower().endswith(".docx") else 0  # docx / doc 97-2003
            doc.SaveAs(os.path.abspath(path_dest), formato)
            return True
        finally:
            try:
                doc.Close(0)  # wdDoNotSaveChanges
            except Exception:
                pass


# ==================== WORKER ====================

class WorkerThread(threading.Thread):
    def __init__(self, arquivos, lingua_destino, cache, queue_msg, stop_event,
                 traduzir_nomes=False, lingua_origem=ORIGEM_AUTO, pular_existentes=False,
                 agrupar=True):
        super().__init__(daemon=True)
        self.arquivos_raw = arquivos
        self.lingua_destino = lingua_destino
        self.lingua_origem = lingua_origem
        self.traduzir_nomes = traduzir_nomes
        self.pular_existentes = pular_existentes
        self.agrupar = agrupar
        self.cache = cache
        self.queue = queue_msg
        self.stop = stop_event
        self.resumo = {"ok": 0, "falha": 0, "pulados": 0, "interrompido": False, "erro": None}

    def run(self):
        if TEM_COM:
            pythoncom.CoInitialize()
        office = SessaoOffice()
        try:
            self._executar(office)
        except Interrompido:
            self.resumo["interrompido"] = True
            self._log("Tradução interrompida pelo usuário.", "warning")
        except ServicoIndisponivel as e:
            self.resumo["erro"] = str(e)
            self._log(f"ERRO: {e}", "error")
        except Exception as e:
            self.resumo["erro"] = str(e)
            self._log(f"ERRO inesperado: {e}", "error")
            self._log(traceback.format_exc(), "error")
        finally:
            office.fechar()
            self.cache.salvar()
            if TEM_COM:
                pythoncom.CoUninitialize()
            self.queue.put(("pronto", self.resumo))

    def _executar(self, office):
        tradutor = Tradutor(self.lingua_destino, self.cache, self._log, self.stop,
                            lingua_origem=self.lingua_origem, agrupar=self.agrupar)

        traducao_nomes = None
        if self.traduzir_nomes:
            self._log("Traduzindo nomes dos arquivos...", "info")
            bases = [os.path.splitext(nome)[0] for _, _, _, nome, _ in self.arquivos_raw]
            traducao_nomes = tradutor.traduzir_batch([b for b in bases if texto_valido(b)])

        sufixo = SUFIXO_IDIOMA.get(self.lingua_destino, "PT_BR")
        arquivos = planejar_destinos(self.arquivos_raw, sufixo, traducao_nomes)

        trad_arq = TradutorArquivos(tradutor, self._log, office, self.stop,
                                    progresso_textos=self._progresso_textos)
        total = len(arquivos)
        for idx, (item_id, path_orig, path_dest, tipo) in enumerate(arquivos, 1):
            if self.stop.is_set():
                raise Interrompido()
            nome = os.path.basename(path_orig)
            self._arquivo_atual = nome
            self._progresso(idx - 1, total, nome)

            if self.pular_existentes and os.path.exists(path_dest):
                self._log(f"[{idx}/{total}] {nome}: já traduzido, pulando.", "info")
                self.resumo["pulados"] += 1
                self.queue.put(("item", item_id, "pulado"))
                continue

            self._log(f"[{idx}/{total}] {nome} ...", "info")
            inicio = time.time()
            try:
                if tipo == "excel":
                    ok = trad_arq.traduzir_excel(path_orig, path_dest)
                else:
                    ok = trad_arq.traduzir_word(path_orig, path_dest)
            except Interrompido:
                raise
            except ServicoIndisponivel:
                self.queue.put(("item", item_id, "falha"))
                raise
            except Exception as e:
                self._log(f"  ERRO: {e}", "error")
                ok = False

            decorrido = time.time() - inicio
            if ok:
                self.resumo["ok"] += 1
                self._log(f"  OK ({decorrido:.0f}s) → {os.path.basename(path_dest)}", "success")
            else:
                self.resumo["falha"] += 1
                self._log(f"  FALHA ({decorrido:.0f}s)", "error")
            self.queue.put(("item", item_id, "ok" if ok else "falha"))
            self._progresso(idx, total, nome)
            gc.collect()

        self._log("Tradução concluída!", "success")

    def _log(self, msg, tipo="info"):
        self.queue.put(("log", msg, tipo))

    def _progresso(self, atual, total, arquivo=""):
        self.queue.put(("progresso", atual, total, arquivo))

    def _progresso_textos(self, feitos, total):
        self.queue.put(("detalhe", f"{getattr(self, '_arquivo_atual', '')} — {feitos}/{total} textos"))


# ==================== INTERFACE GRÁFICA ====================

class TradutorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Tradutor de arquivos office")
        self.root.geometry("980x760")
        self.root.minsize(780, 640)
        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar)

        self._icone_caminho = self._localizar_icone()
        if self._icone_caminho.exists():
            try:
                self.root.iconbitmap(default=str(self._icone_caminho))
            except Exception:
                pass

        self.cache = CacheManager()
        self.queue_app = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None
        self._fechando = False
        self.arquivos_info = []
        self.arquivos_check = {}

        self._config = self._carregar_config()
        self._limpar_atualizacao_antiga()

        self._criar_menu()
        self._criar_widgets()
        self._ajustar_tema()
        self._atualizar_info_cache()
        self._poll_queue()

    def _localizar_icone(self):
        embutido = Path(getattr(sys, "_MEIPASS", DIR_APP)) / "icone_app.ico"
        if embutido.exists():
            return embutido
        caminho = DIR_DADOS / "icone_app.ico"
        if not caminho.exists():
            try:
                self._gerar_icone(caminho)
            except OSError:
                pass
        return caminho

    @staticmethod
    def _gerar_icone(destino):
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
                        painter[(x, y)] = BRANCO
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

        with open(destino, 'wb') as f:
            f.write(struct.pack('<HHH', 0, 1, 1))
            f.write(struct.pack('<BBBBHHII', W, H, 0, 0, 1, 32, bmp_size, off))
            f.write(struct.pack('<IiiHHIIiiII', 40, W, H * 2, 1, 32, 0, len(bmp_padded), 0, 0, 0, 0))
            f.write(bmp_padded)
            f.write(and_mask)

    # ---------- Configuração ----------

    @staticmethod
    def _carregar_config():
        try:
            cfg = json.loads(ARQ_CONFIG.read_text("utf-8"))
            return cfg if isinstance(cfg, dict) else {}
        except (OSError, ValueError):
            return {}

    def _salvar_config(self):
        cfg = {
            "ultimo_dir": self.var_pasta.get().strip(),
            "idioma": self.var_lingua.get(),
            "origem": self.var_origem.get(),
            "excel": self.var_excel.get(),
            "word": self.var_word.get(),
            "traduzir_nomes": self.var_traduzir_nomes.get(),
            "pular_existentes": self.var_pular.get(),
            "agrupar": self.var_agrupar.get(),
        }
        try:
            ARQ_CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), "utf-8")
        except OSError:
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
        except tk.TclError:
            pass
        familia = fonte_padrao.actual("family")
        style.configure("Accent.TButton", font=(familia, 10, "bold"))
        style.configure("Treeview", rowheight=28, font=(familia, 10))
        style.configure("Treeview.Heading", font=(familia, 10, "bold"))
        style.configure("TLabelframe.Label", font=(familia, 10, "bold"))
        style.configure("TButton", font=(familia, 10), padding=(8, 4))
        style.configure("TCheckbutton", font=(familia, 10))
        style.configure("TLabel", font=(familia, 10))
        style.configure("TEntry", font=(familia, 10))

    # ---------- Menu / janelas auxiliares ----------

    def _criar_menu(self):
        barra = tk.Menu(self.root)
        menu_ajuda = tk.Menu(barra, tearoff=0)
        menu_ajuda.add_command(label="Como usar", command=self._mostrar_ajuda)
        menu_ajuda.add_separator()
        menu_ajuda.add_command(label="Verificar atualizações", command=self._verificar_atualizacao)
        menu_ajuda.add_separator()
        menu_ajuda.add_command(label="Sobre", command=self._mostrar_sobre)
        barra.add_cascade(label="Ajuda", menu=menu_ajuda)
        self.root.configure(menu=barra)

    def _nova_janela(self, titulo, geometria):
        top = tk.Toplevel(self.root)
        top.title(titulo)
        top.geometry(geometria)
        top.minsize(400, 320)
        top.transient(self.root)
        top.grab_set()
        if self._icone_caminho.exists():
            try:
                top.iconbitmap(default=str(self._icone_caminho))
            except Exception:
                pass
        return top

    def _mostrar_ajuda(self):
        top = self._nova_janela("Como usar", "560x520")

        ttk.Button(top, text="Fechar", command=top.destroy).pack(side="bottom", pady=(0, 8))
        quadro = ttk.Frame(top)
        quadro.pack(fill="both", expand=True, padx=8, pady=8)
        txt = tk.Text(quadro, wrap="word", font=("Segoe UI", 10),
                      relief="sunken", borderwidth=2, padx=12, pady=12)
        scroll = ttk.Scrollbar(quadro, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)

        txt.insert("end", "COMO USAR O TRADUTOR DE ARQUIVOS OFFICE\n\n", "titulo")
        txt.insert("end", "1. Selecione a pasta raiz com os arquivos clicando em 📁.\n")
        txt.insert("end", "2. Escolha os tipos de arquivo desejados (Excel, Word).\n")
        txt.insert("end", "3. Selecione o idioma de origem (ou deixe em detecção automática) "
                          "e o idioma de destino.\n")
        txt.insert("end", '4. Marque "Traduzir nomes" se quiser traduzir também os nomes dos arquivos.\n')
        txt.insert("end", "5. Clique em 🔍 Escanear para listar os arquivos encontrados.\n")
        txt.insert("end", "6. Clique na coluna de seleção (☐/☑) para marcar/desmarcar arquivos.\n")
        txt.insert("end", "   Use a tecla Espaço para alternar a seleção do item em foco.\n")
        txt.insert("end", "7. Clique em ▶ Iniciar Tradução para começar.\n\n")
        txt.insert("end", "Os arquivos traduzidos são salvos com o sufixo do idioma "
                          "(ex.: _PT_BR) ao lado do original. Arquivos .xls são salvos como .xlsx.\n\n")
        txt.insert("end", "Use ⏹ Parar para interromper; o arquivo em andamento não é salvo. "
                          'Com "Pular já traduzidos" marcado, uma nova execução continua de onde parou.\n\n')
        txt.insert("end", "As traduções ficam em cache (por idioma) e são reaproveitadas. "
                          "O botão Limpar Cache força uma retradução completa.\n\n")
        txt.insert("end", "Parágrafos do Word que contêm imagens, campos ou notas de rodapé são "
                          "mantidos no original para não perder esses elementos (o log informa).\n\n")
        txt.insert("end", "DICA: arquivos gerados pelo programa (sufixo de idioma ou _traduzido "
                          "no nome) são ignorados automaticamente no escaneamento.", "dica")

        txt.tag_configure("titulo", font=("Segoe UI", 12, "bold"), spacing3=8)
        txt.tag_configure("dica", foreground="#0044aa", font=("Segoe UI", 10, "italic"))
        txt.configure(state="disabled")

    def _mostrar_sobre(self):
        msg = (
            f"Tradutor de Arquivos Office v{VERSAO}\n\n"
            "Traduz conteúdo de arquivos Excel e Word\n"
            "mantendo a formatação original (cores, layout, fontes).\n\n"
            "Tecnologias: Python, Tkinter, deep-translator\n"
            "Padrões: WCAG AA, ISO 9241-115/161, HIG\n\n"
            "Desenvolvido por Autolinx Automação\n"
            "http://autolinx.me\n\n"
            "© 2026"
        )
        messagebox.showinfo("Sobre", msg)

    # ---------- Atualização ----------

    def _limpar_atualizacao_antiga(self):
        if getattr(sys, "frozen", False):
            try:
                (DIR_APP / "TradutorOffice.old.exe").unlink()
            except OSError:
                pass

    def _verificar_atualizacao(self):
        self._log("Verificando atualizações...", "info")

        def tarefa():
            try:
                req = urllib.request.Request(REPO_API, headers={"User-Agent": "TradutorOffice"})
                with urllib.request.urlopen(req, timeout=10) as resp:
                    self.queue_app.put(("atualizacao", json.loads(resp.read()), None))
            except Exception as e:
                self.queue_app.put(("atualizacao", None, e))

        threading.Thread(target=tarefa, daemon=True).start()

    def _mostrar_atualizacao(self, dados, erro):
        if erro is not None:
            messagebox.showerror("Erro", f"Falha ao verificar atualização:\n{erro}")
            return
        ultima = (dados.get("tag_name") or "").lstrip("vV")
        if not versao_tupla(ultima):
            messagebox.showwarning("Atualização", "Não foi possível identificar a versão remota.")
            return
        if versao_tupla(ultima) <= versao_tupla(VERSAO):
            messagebox.showinfo("Atualizado", f"Você já está na versão mais recente (v{VERSAO}).")
            return

        corpo = dados.get("body") or "Sem detalhes da versão."
        nome_release = dados.get("name") or f"v{ultima}"
        asset_url = None
        for asset in dados.get("assets", []):
            if asset.get("name", "").lower().endswith(".exe"):
                asset_url = asset.get("browser_download_url")
                break

        top = self._nova_janela(f"Atualização disponível — {nome_release}", "560x480")
        frame = ttk.Frame(top, padding=12)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=f"Versão atual: v{VERSAO}    →    {nome_release}",
                  font=(None, 11, "bold")).pack(anchor="w")
        ttk.Label(frame, text="O que há de novo:",
                  font=(None, 10, "bold")).pack(anchor="w", pady=(10, 2))

        row_btn = ttk.Frame(frame)
        row_btn.pack(side="bottom", fill="x", pady=(10, 0))
        quadro = ttk.Frame(frame)
        quadro.pack(fill="both", expand=True)
        txt = tk.Text(quadro, wrap="word", font=("Segoe UI", 10),
                      relief="sunken", borderwidth=2, height=12)
        scroll = ttk.Scrollbar(quadro, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)
        txt.insert("end", corpo)
        txt.configure(state="disabled")

        if asset_url and getattr(sys, "frozen", False):
            ttk.Button(row_btn, text="  ⬇ Baixar e Atualizar  ",
                       command=lambda: self._baixar_atualizar(asset_url, top)
                       ).pack(side="left", padx=3)
        elif asset_url:
            ttk.Label(row_btn, text="(modo desenvolvimento — abrindo navegador)",
                      foreground="#555").pack(side="left", padx=3)
            ttk.Button(row_btn, text="  Abrir Release  ",
                       command=lambda: webbrowser.open(REPO_URL)
                       ).pack(side="left", padx=3)
        else:
            ttk.Label(row_btn, text="Nenhum instalador disponível neste release.",
                      foreground="#555").pack(side="left", padx=3)

        ttk.Button(row_btn, text="Cancelar", command=top.destroy).pack(side="right", padx=3)

    def _baixar_atualizar(self, url, janela):
        janela.destroy()
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("Atualizar", "Aguarde o fim da tradução antes de atualizar.")
            return
        host = (urlparse(url).hostname or "").lower()
        if urlparse(url).scheme != "https" or not (host == "github.com" or host.endswith(".github.com")):
            messagebox.showerror("Erro", f"Endereço de download inesperado:\n{url}")
            return

        exe_atual = Path(sys.executable)
        exe_novo = DIR_APP / "TradutorOffice.novo.exe"
        exe_old = DIR_APP / "TradutorOffice.old.exe"
        top_prog = None
        try:
            top_prog = self._nova_janela("Baixando atualização...", "400x120")
            ttk.Label(top_prog, text="Baixando nova versão...",
                      font=(None, 10, "bold")).pack(pady=(12, 4))
            prog = ttk.Progressbar(top_prog, mode="determinate")
            prog.pack(fill="x", padx=20, pady=4)
            lbl = ttk.Label(top_prog, text="")
            lbl.pack()
            top_prog.update()

            def andamento(blocos, tam_bloco, total):
                baixado = blocos * tam_bloco
                if total > 0:
                    prog["value"] = min(100, baixado * 100 / total)
                lbl.configure(text=f"{baixado // 1024} / {max(total, 0) // 1024} KB")
                top_prog.update()

            urllib.request.urlretrieve(url, exe_novo, reporthook=andamento)
            top_prog.destroy()
            top_prog = None

            if not exe_novo.exists() or exe_novo.stat().st_size == 0:
                messagebox.showerror("Erro", "Falha ao baixar o arquivo.")
                return

            if messagebox.askyesno("Atualizar",
                                   "Download concluído. Deseja instalar agora?\n"
                                   "O aplicativo será fechado e reaberto automaticamente."):
                caminho_bat = DIR_APP / "atualizar.bat"
                # chcp 65001: o cmd lê o .bat em UTF-8 (caminhos com acentos)
                caminho_bat.write_text(
                    "@echo off\r\n"
                    "chcp 65001 >nul\r\n"
                    "timeout /t 2 /nobreak >nul\r\n"
                    f'move /y "{exe_atual}" "{exe_old}" >nul 2>&1\r\n'
                    f'move /y "{exe_novo}" "{exe_atual}" >nul 2>&1\r\n'
                    f'start "" "{exe_atual}"\r\n'
                    'del "%~f0"\r\n',
                    encoding="utf-8",
                )
                os.startfile(caminho_bat)
                self._salvar_config()
                self.root.quit()

        except Exception as e:
            if top_prog is not None:
                top_prog.destroy()
            try:
                exe_novo.unlink()
            except OSError:
                pass
            messagebox.showerror("Erro", f"Falha ao atualizar:\n{e}")

    # ---------- Widgets ----------

    def _criar_widgets(self):
        cfg = self._config
        idioma = cfg.get("idioma") if cfg.get("idioma") in LINGUAS else "Português (BR)"
        origem = cfg.get("origem") if cfg.get("origem") in LINGUAS_ORIGEM else ORIGEM_AUTO

        # ===== VARIÁVEIS =====
        self.var_pasta = tk.StringVar(value=cfg.get("ultimo_dir", ""))
        self.var_excel = tk.BooleanVar(value=cfg.get("excel", True))
        self.var_word = tk.BooleanVar(value=cfg.get("word", True))
        self.var_lingua = tk.StringVar(value=idioma)
        self.var_origem = tk.StringVar(value=origem)
        self.var_traduzir_nomes = tk.BooleanVar(value=cfg.get("traduzir_nomes", False))
        self.var_pular = tk.BooleanVar(value=cfg.get("pular_existentes", False))
        self.var_agrupar = tk.BooleanVar(value=cfg.get("agrupar", True))

        # ===== FRAME TOPO =====
        frame_topo = ttk.LabelFrame(self.root, text="Configuração", padding=12)
        frame_topo.pack(fill="x", padx=12, pady=(12, 6))

        # Pasta + Escanear
        row_pasta = ttk.Frame(frame_topo)
        row_pasta.pack(fill="x", pady=4)
        ttk.Label(row_pasta, text="Pasta raiz:", width=16).pack(side="left")
        self.entry_pasta = ttk.Entry(row_pasta, textvariable=self.var_pasta)
        self.entry_pasta.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.entry_pasta.bind("<Return>", lambda e: self._escanear())
        self.btn_pasta = ttk.Button(row_pasta, text="  📁  ", command=self._selecionar_pasta, width=4)
        self.btn_pasta.pack(side="left")
        self.btn_escanear = ttk.Button(row_pasta, text="  🔍 Escanear  ", command=self._escanear)
        self.btn_escanear.pack(side="left", padx=(6, 0))

        # Tipos
        row_tipos = ttk.Frame(frame_topo)
        row_tipos.pack(fill="x", pady=4)
        ttk.Label(row_tipos, text="Tipos de arquivo:", width=16).pack(side="left")
        self.chk_excel = ttk.Checkbutton(row_tipos, text="Excel (.xls/.xlsx)", variable=self.var_excel)
        self.chk_excel.pack(side="left", padx=(0, 12))
        self.chk_word = ttk.Checkbutton(row_tipos, text="Word (.doc/.docx)", variable=self.var_word)
        self.chk_word.pack(side="left", padx=(0, 12))

        # Idiomas
        row_idioma = ttk.Frame(frame_topo)
        row_idioma.pack(fill="x", pady=4)
        ttk.Label(row_idioma, text="Idioma origem:", width=16).pack(side="left")
        ttk.Combobox(
            row_idioma, textvariable=self.var_origem,
            values=LINGUAS_ORIGEM, state="readonly", width=24,
        ).pack(side="left", padx=(0, 12))
        ttk.Label(row_idioma, text="→  Destino:").pack(side="left", padx=(0, 6))
        ttk.Combobox(
            row_idioma, textvariable=self.var_lingua,
            values=LINGUAS, state="readonly", width=20,
        ).pack(side="left")

        # Opções
        row_opcoes = ttk.Frame(frame_topo)
        row_opcoes.pack(fill="x", pady=4)
        ttk.Label(row_opcoes, text="Opções:", width=16).pack(side="left")
        self.chk_trad_nomes = ttk.Checkbutton(
            row_opcoes, text="Traduzir nomes", variable=self.var_traduzir_nomes)
        self.chk_trad_nomes.pack(side="left", padx=(0, 12))
        self.chk_pular = ttk.Checkbutton(
            row_opcoes, text="Pular já traduzidos", variable=self.var_pular)
        self.chk_pular.pack(side="left", padx=(0, 12))
        self.chk_agrupar = ttk.Checkbutton(
            row_opcoes, text="Agrupar textos (menos requisições)", variable=self.var_agrupar)
        self.chk_agrupar.pack(side="left")

        # ===== FRAME BOTÕES (empacotado antes da lista para nunca sumir ao redimensionar) =====
        frame_botoes = ttk.Frame(self.root, padding=6)
        frame_botoes.pack(side="bottom", fill="x", padx=12, pady=(0, 12))

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

        self.btn_limpar_cache = ttk.Button(
            frame_botoes, text=" Limpar Cache ",
            command=self._limpar_cache,
        )
        self.btn_limpar_cache.pack(side="right", padx=3)
        self.lbl_cache = ttk.Label(frame_botoes, text="", foreground="#555")
        self.lbl_cache.pack(side="right", padx=6)

        self.lbl_status = ttk.Label(frame_botoes, text="Pronto", foreground="#444", font=(None, 10, "bold"))
        self.lbl_status.pack(side="right", padx=16)

        # ===== FRAME LOG =====
        frame_log = ttk.LabelFrame(self.root, text="Log", padding=6)
        frame_log.pack(side="bottom", fill="x", padx=12, pady=6)

        self.txt_log = tk.Text(frame_log, height=6, wrap="word", state="disabled",
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

        # ===== FRAME PROGRESSO =====
        frame_prog = ttk.LabelFrame(self.root, text="Progresso", padding=8)
        frame_prog.pack(side="bottom", fill="x", padx=12, pady=6)

        row_prog = ttk.Frame(frame_prog)
        row_prog.pack(fill="x")
        self.progresso_bar = ttk.Progressbar(row_prog, mode="determinate")
        self.progresso_bar.pack(side="left", fill="x", expand=True)
        self.lbl_progresso = ttk.Label(row_prog, text="0 / 0", width=10, anchor="center")
        self.lbl_progresso.pack(side="left", padx=(12, 0))
        self.lbl_arquivo_atual = ttk.Label(row_prog, text="", width=44, anchor="e")
        self.lbl_arquivo_atual.pack(side="left", padx=(12, 0))

        # ===== FRAME LISTA =====
        frame_lista = ttk.LabelFrame(self.root, text="Arquivos encontrados", padding=6)
        frame_lista.pack(fill="both", expand=True, padx=12, pady=6)

        row_sel = ttk.Frame(frame_lista)
        row_sel.pack(side="top", fill="x", pady=(0, 6))
        self.lbl_selecao = ttk.Label(row_sel, text="", foreground="#444")
        self.lbl_selecao.pack(side="left")
        self.btn_limpar_sel = ttk.Button(row_sel, text=" Desmarcar Todos ", command=self._limpar_selecao)
        self.btn_limpar_sel.pack(side="right", padx=3)
        self.btn_sel_todos = ttk.Button(row_sel, text=" Selecionar Todos ", command=self._selecionar_todos)
        self.btn_sel_todos.pack(side="right", padx=3)

        cols = ("check", "arquivo", "pasta", "tamanho")
        self.tree = ttk.Treeview(
            frame_lista, columns=cols, show="headings",
            selectmode="none", height=6,
        )
        self.tree.heading("check", text="☑", command=self._alternar_todos)
        self.tree.heading("arquivo", text="Arquivo")
        self.tree.heading("pasta", text="Pasta")
        self.tree.heading("tamanho", text="Tamanho")
        self.tree.column("check", width=36, anchor="center", stretch=False)
        self.tree.column("arquivo", width=300, anchor="w")
        self.tree.column("pasta", width=300, anchor="w")
        self.tree.column("tamanho", width=90, anchor="e", stretch=False)
        self.tree.bind("<ButtonRelease-1>", self._toggle_check)
        self.tree.bind("<space>", self._toggle_check_key)
        self.tree.tag_configure("ok", background="#D4EDDA")
        self.tree.tag_configure("falha", background="#F8D7DA")
        self.tree.tag_configure("pulado", background="#E2E3E5")

        scroll_y = ttk.Scrollbar(frame_lista, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

    # ---------- Ações ----------

    def _selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecionar pasta raiz",
                                        initialdir=self.var_pasta.get() or None)
        if pasta:
            self.var_pasta.set(pasta)
            self._salvar_config()

    def _escanear(self):
        if self.worker and self.worker.is_alive():
            return
        pasta = self.var_pasta.get().strip()
        if not pasta or not os.path.isdir(pasta):
            messagebox.showwarning("Aviso", "Selecione uma pasta válida.")
            return
        self._salvar_config()

        self._log("Escaneando arquivos...", "info")
        self.arquivos_info.clear()
        self.arquivos_check.clear()
        self.tree.delete(*self.tree.get_children())

        sufixo = SUFIXO_IDIOMA.get(self.var_lingua.get(), "PT_BR")
        try:
            encontrados = escanear_pasta(pasta, self.var_excel.get(), self.var_word.get(), sufixo)
        except OSError as e:
            messagebox.showerror("Erro", f"Falha ao escanear a pasta:\n{e}")
            return

        for path, raiz, nome, tipo in encontrados:
            nome_pasta = os.path.relpath(raiz, pasta)
            if nome_pasta == ".":
                nome_pasta = ""
            try:
                str_tam = formatar_tamanho(os.path.getsize(path))
            except OSError:
                str_tam = "?"
            item_id = self.tree.insert("", "end", values=("☑", nome, nome_pasta, str_tam))
            self.arquivos_check[item_id] = True
            self.arquivos_info.append((item_id, path, raiz, nome, tipo))

        self._atualizar_contagem()
        filhos = self.tree.get_children()
        if filhos:
            self.tree.focus(filhos[0])
        n = len(encontrados)
        self._log(f"{n} arquivo(s) encontrado(s).", "success" if n else "warning")
        if not n:
            self._log("Nenhum arquivo corresponde aos filtros selecionados.", "warning")

    def _marcar(self, item, valor):
        self.arquivos_check[item] = valor
        self.tree.set(item, "check", "☑" if valor else "☐")
        self._atualizar_contagem()

    def _atualizar_contagem(self):
        total = len(self.arquivos_check)
        marcados = sum(self.arquivos_check.values())
        self.lbl_selecao["text"] = f"{marcados} de {total} selecionado(s)" if total else ""

    def _toggle_check(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        item = self.tree.identify_row(event.y)
        if not item or item not in self.arquivos_check:
            return
        self.tree.focus(item)
        if self.tree.identify_column(event.x) != "#1":
            return
        self._marcar(item, not self.arquivos_check[item])

    def _toggle_check_key(self, event):
        item = self.tree.focus()
        if item and item in self.arquivos_check:
            self._marcar(item, not self.arquivos_check[item])
            return "break"

    def _selecionar_todos(self):
        for item in self.arquivos_check:
            self._marcar(item, True)

    def _limpar_selecao(self):
        for item in self.arquivos_check:
            self._marcar(item, False)

    def _alternar_todos(self):
        if self.arquivos_check and all(self.arquivos_check.values()):
            self._limpar_selecao()
        else:
            self._selecionar_todos()

    def _definir_executando(self, executando):
        normal = "disabled" if executando else "normal"
        for w in (self.btn_iniciar, self.btn_escanear, self.btn_limpar_cache,
                  self.btn_sel_todos, self.btn_limpar_sel):
            w["state"] = normal
        self.btn_parar["state"] = "normal" if executando else "disabled"

    def _iniciar(self):
        if self.worker and self.worker.is_alive():
            return
        if not self.arquivos_info:
            messagebox.showwarning("Aviso", "Escaneie os arquivos primeiro.")
            return

        selecionados = [info for info in self.arquivos_info if self.arquivos_check.get(info[0])]
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos um arquivo.")
            return

        origem, destino = self.var_origem.get(), self.var_lingua.get()
        if COD_LINGUAS.get(origem) == COD_LINGUAS.get(destino):
            messagebox.showwarning("Aviso", "O idioma de origem e o de destino são iguais.")
            return
        precisa_com = any(tipo == "word" or nome.lower().endswith(".xls")
                          for _, _, _, nome, tipo in selecionados)
        if precisa_com and not TEM_COM:
            if not messagebox.askyesno(
                    "Microsoft Office indisponível",
                    "Arquivos .xls, .doc e .docx exigem Windows com Microsoft Office e pywin32.\n"
                    "Eles falharão. Continuar mesmo assim?"):
                return

        self._salvar_config()
        self.stop_event.clear()
        for item_id, *_ in selecionados:
            self.tree.item(item_id, tags=())

        self.progresso_bar["value"] = 0
        self.lbl_progresso["text"] = f"0 / {len(selecionados)}"
        self.lbl_arquivo_atual["text"] = ""
        self._definir_executando(True)
        self.lbl_status["text"] = "Traduzindo..."
        self.lbl_status["foreground"] = "#0059a0"

        self.worker = WorkerThread(
            selecionados, destino, self.cache, self.queue_app, self.stop_event,
            traduzir_nomes=self.var_traduzir_nomes.get(), lingua_origem=origem,
            pular_existentes=self.var_pular.get(), agrupar=self.var_agrupar.get(),
        )
        self.worker.start()

    def _parar(self):
        # Parada cooperativa: o worker fecha o Excel/Word e grava o cache antes de terminar
        if self.worker and self.worker.is_alive():
            self.stop_event.set()
            self._log("Parada solicitada; finalizando a operação atual...", "warning")
            self.lbl_status["text"] = "Parando..."
            self.lbl_status["foreground"] = "#7a5c00"
        self.btn_parar["state"] = "disabled"

    def _finalizar(self, resumo):
        self._definir_executando(False)
        self.lbl_arquivo_atual["text"] = ""
        self._atualizar_info_cache()
        if self._fechando:
            return
        partes = [f"{resumo['ok']} traduzido(s)"]
        if resumo["falha"]:
            partes.append(f"{resumo['falha']} com falha")
        if resumo["pulados"]:
            partes.append(f"{resumo['pulados']} pulado(s)")
        texto = ", ".join(partes)

        if resumo["interrompido"]:
            self.lbl_status["text"] = "Interrompido"
            self.lbl_status["foreground"] = "#7a5c00"
        elif resumo["erro"]:
            self.lbl_status["text"] = "Erro"
            self.lbl_status["foreground"] = "#b30000"
            messagebox.showerror("Erro", f"{resumo['erro']}\n\n{texto}.")
        elif resumo["falha"]:
            self.lbl_status["text"] = "Concluído com falhas"
            self.lbl_status["foreground"] = "#b30000"
            messagebox.showwarning("Concluído", f"Tradução finalizada: {texto}.\nVeja o log para detalhes.")
        else:
            self.lbl_status["text"] = "Concluído"
            self.lbl_status["foreground"] = "#0d6b0d"
            messagebox.showinfo("Concluído", f"Tradução finalizada com sucesso: {texto}.")

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
                    self.lbl_arquivo_atual["text"] = arquivo
                elif tipo == "detalhe":
                    self.lbl_arquivo_atual["text"] = msg[1]
                elif tipo == "item":
                    _, item_id, estado = msg
                    if self.tree.exists(item_id):
                        self.tree.item(item_id, tags=(estado,))
                elif tipo == "pronto":
                    self._finalizar(msg[1])
                elif tipo == "atualizacao":
                    self._mostrar_atualizacao(msg[1], msg[2])
        except queue.Empty:
            pass

        self.root.after(100, self._poll_queue)

    def _log(self, texto, tipo="info"):
        self.txt_log.configure(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{timestamp}] {texto}\n", tipo)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _atualizar_info_cache(self):
        self.lbl_cache["text"] = f"Cache: {self.cache.total()} traduções"

    def _limpar_cache(self):
        if self.worker and self.worker.is_alive():
            return
        if messagebox.askyesno("Limpar Cache", "Limpar o cache de traduções de todos os idiomas?"):
            self.cache.limpar()
            self._atualizar_info_cache()
            self._log("Cache limpo.", "warning")

    def _ao_fechar(self):
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("Sair", "Há uma tradução em andamento. Interromper e sair?"):
                return
            self._fechando = True
            self.stop_event.set()
            self.lbl_status["text"] = "Encerrando..."
            self._aguardar_e_fechar(0)
            return
        self._salvar_config()
        self.root.destroy()

    def _aguardar_e_fechar(self, tentativas):
        # Espera o worker fechar o Excel/Word (até ~20 s) para não deixar processos órfãos
        if self.worker.is_alive() and tentativas < 200:
            self.root.after(100, lambda: self._aguardar_e_fechar(tentativas + 1))
            return
        self._salvar_config()
        self.root.destroy()

    def executar(self):
        self.root.mainloop()


# ==================== MAIN ====================

if __name__ == "__main__":
    app = TradutorApp()
    app.executar()
