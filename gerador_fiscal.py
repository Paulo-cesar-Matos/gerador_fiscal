# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingTypeStubs=false, reportAttributeAccessIssue=false
import os
import re
import unicodedata
from datetime import datetime
from typing import Any, Dict, List, Literal, Optional, Tuple
import numpy as np # type: ignore
import pandas as pd # pyright: ignore[reportMissingModuleSource]
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pyperclip # pyright: ignore[reportMissingModuleSource]
except ImportError:
    pyperclip = None


def normalize_header_text(text: Any) -> str:
    """Normaliza nomes de colunas removendo acentos, pontuações e símbolos ordinais."""
    if text is None:
        return ""
    s = unicodedata.normalize("NFKD", str(text))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.replace("º", " ").replace("°", " ").replace("ª", " ")
    s = re.sub(r"[^a-zA-Z0-9\s]", " ", s)
    return " ".join(s.split()).upper()


def format_cell_value(val: Any) -> str:
    """Formata valores de células do Excel, evitando o bug de '.0' em números inteiros."""
    if val is None or pd.isna(val):
        return "-"

    if isinstance(val, (int, np.integer)):
        return str(val)

    if isinstance(val, (float, np.floating)):
        if np.isnan(val):
            return "-"
        if val.is_integer():
            return str(int(val))
        return str(val)

    s = str(val).strip()
    if s.lower() in ("", "nan", "none", "null", "<na>", "-"):
        return "-"

    # Converte números em formato texto terminados em .0 (ex: "6495591.0")
    if re.match(r"^-?\d+\.0+$", s):
        return s.split(".")[0]

    return s


def format_date_str(val: Any) -> str:
    """Formata datas para o padrão DD/MM/AAAA."""
    if val is None or pd.isna(val):
        return ""

    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%d/%m/%Y")

    s = str(val).strip()
    if not s or s.lower() in ("", "nan", "none", "null", "<na>", "-", "."):
        return ""

    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass

    try:
        f = float(s)
        if 30000 < f < 60000:
            dt = pd.to_datetime(f, unit="D", origin="1899-12-30")
            if isinstance(dt, (datetime, pd.Timestamp)): # pyright: ignore[reportUnnecessaryIsInstance]
                return dt.strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        pass

    return s.split()[0] if " " in s else s


def parse_date_key(date_str: str) -> datetime:
    """Converte string DD/MM/AAAA para datetime para ordenação correta."""
    try:
        return datetime.strptime(date_str.strip(), "%d/%m/%Y")
    except Exception:
        return datetime.min


def extract_maps_link(val: Any) -> str:
    """Extrai URL do Google Maps ou coordenadas geográficas válidas."""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if s in ("", ".", "-"):
        return ""

    match = re.search(r"https?://[^\s]+", s)
    if match:
        return match.group(0).rstrip(".,;")

    # Coordenadas em formato "lat, long"
    if re.match(r"^[-+]?\d+\.\d+\s*,\s*[-+]?\d+\.\d+$", s):
        return s

    return ""


def get_time_greeting() -> str:
    """Retorna saudação apropriada com base na hora do dia."""
    hora = datetime.now().hour
    if hora < 12:
        return "Bom dia"
    elif hora < 18:
        return "Boa tarde"
    return "Boa noite"


class AppFiscal:
    def __init__(self, root: tk.Tk) -> None:
        self.root: tk.Tk = root
        self.root.title("Gerador de Relatórios para Fiscais")
        self.root.geometry("900x740")
        self.root.minsize(840, 640)

        self.df: Optional[pd.DataFrame] = None
        self.fiscais_list: List[str] = []
        self.fiscais_filtrados: List[str] = []
        self.datas_fiscal: List[Tuple[str, int]] = []
        self.datas_fiscal_map: Dict[str, int] = {}
        self.var_ordem_datas: tk.StringVar = tk.StringVar(value="decrescente")
        self.total_ordens_fiscal_atual: int = 0
        self.data_selecionada: Optional[str] = None
        self.col_map: Dict[str, str] = {}
        self.caminho_arquivo: str = ""
        self.aba_ativa: str = ""

        # Paleta de Cores da Interface
        self.bg_color: str = "#f8fafc"
        self.header_bg: str = "#0f172a"
        self.header_fg: str = "#ffffff"
        self.accent_color: str = "#2563eb"
        self.accent_hover: str = "#1d4ed8"
        self.success_color: str = "#059669"
        self.success_hover: str = "#047857"

        self.root.configure(bg=self.bg_color)
        self._setup_ui()

    def _setup_ui(self) -> None:
        # Cabeçalho Superior
        top_frame = tk.Frame(self.root, bg=self.header_bg, height=65)
        top_frame.pack(fill="x", side="top")

        lbl_container = tk.Frame(top_frame, bg=self.header_bg)
        lbl_container.pack(side="left", padx=20, pady=12)

        lbl_title = tk.Label(
            lbl_container,
            text="📋 Gerador de Mensagens de Fiscalização",
            font=("Segoe UI", 13, "bold"),
            bg=self.header_bg,
            fg=self.header_fg,
        )
        lbl_title.pack(anchor="w")

        self.lbl_substatus = tk.Label(
            lbl_container,
            text="Nenhuma planilha importada",
            font=("Segoe UI", 9),
            bg=self.header_bg,
            fg="#94a3b8",
        )
        self.lbl_substatus.pack(anchor="w")

        btn_import = tk.Button(
            top_frame,
            text="📁 Importar Planilha Excel",
            command=self.importar_planilha,
            font=("Segoe UI", 10, "bold"),
            bg=self.success_color,
            fg="white",
            activebackground=self.success_hover,
            activeforeground="white",
            relief="flat",
            padx=16,
            pady=8,
            cursor="hand2",
        )
        btn_import.pack(side="right", padx=20, pady=12)

        # Container Principal
        main_container = tk.Frame(self.root, bg=self.bg_color)
        main_container.pack(fill="both", expand=True, padx=15, pady=(12, 5))

        # Painel Esquerdo: Menu de Fiscais e Datas
        left_panel = tk.LabelFrame(
            main_container,
            text=" Mandar para: ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color,
            fg="#334155",
            padx=10,
            pady=10,
            width=280,
        )
        left_panel.pack(side="left", fill="y", padx=(0, 10))
        left_panel.pack_propagate(False)

        # Caixa de Pesquisa de Fiscais
        search_frame = tk.Frame(left_panel, bg=self.bg_color)
        search_frame.pack(fill="x", pady=(0, 6))

        self.txt_busca = tk.Entry(
            search_frame,
            font=("Segoe UI", 10),
            relief="solid",
            bd=1,
            fg="#64748b",
        )
        self.txt_busca.insert(0, "🔍 Buscar fiscal...")
        self.txt_busca.pack(side="left", fill="x", expand=True)
        self.txt_busca.bind("<FocusIn>", self._on_search_focus_in)
        self.txt_busca.bind("<FocusOut>", self._on_search_focus_out)
        self.txt_busca.bind("<KeyRelease>", self._on_search_change)

        btn_clear_search = tk.Button(
            search_frame,
            text="✖",
            command=self._clear_search,
            font=("Segoe UI", 8),
            relief="flat",
            bg="#e2e8f0",
            fg="#64748b",
            padx=5,
            cursor="hand2",
        )
        btn_clear_search.pack(side="right", padx=(2, 0))

        # Lista de Fiscais com Scrollbars
        list_frame = tk.Frame(left_panel, bg=self.bg_color)
        list_frame.pack(fill="both", expand=True)

        scrollbar_y = tk.Scrollbar(list_frame, orient="vertical")
        scrollbar_x = tk.Scrollbar(list_frame, orient="horizontal")

        self.lst_fiscais = tk.Listbox(
            list_frame,
            font=("Segoe UI", 10),
            selectmode=tk.SINGLE,
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
            exportselection=False,
            height=10,
            relief="solid",
            bd=1,
            activestyle="none",
        )
        scrollbar_y.config(command=self.lst_fiscais.yview)
        scrollbar_x.config(command=self.lst_fiscais.xview)

        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        self.lst_fiscais.pack(side="left", fill="both", expand=True)
        self.lst_fiscais.bind("<<ListboxSelect>>", self.on_fiscal_select)

        # Seção de Datas com Opção Crescente e Decrescente
        date_header_frame = tk.Frame(left_panel, bg=self.bg_color)
        date_header_frame.pack(fill="x", pady=(10, 3))

        lbl_datas = tk.Label(
            date_header_frame,
            text="📅 Data:",
            font=("Segoe UI", 9, "bold"),
            bg=self.bg_color,
            fg="#1e293b",
        )
        lbl_datas.pack(side="left")

        rb_dec = tk.Radiobutton(
            date_header_frame,
            text="⬇ Decres.",
            variable=self.var_ordem_datas,
            value="decrescente",
            command=self.aplicar_ordem_datas,
            font=("Segoe UI", 8),
            bg=self.bg_color,
            activebackground=self.bg_color,
            cursor="hand2",
        )
        rb_dec.pack(side="right", padx=(1, 0))

        rb_cres = tk.Radiobutton(
            date_header_frame,
            text="⬆ Cres.",
            variable=self.var_ordem_datas,
            value="crescente",
            command=self.aplicar_ordem_datas,
            font=("Segoe UI", 8),
            bg=self.bg_color,
            activebackground=self.bg_color,
            cursor="hand2",
        )
        rb_cres.pack(side="right", padx=(1, 0))

        date_frame = tk.Frame(left_panel, bg=self.bg_color)
        date_frame.pack(fill="x", pady=(0, 6))

        scrollbar_datas = tk.Scrollbar(date_frame, orient="vertical")
        self.lst_datas = tk.Listbox(
            date_frame,
            font=("Segoe UI", 9),
            selectmode=tk.SINGLE,
            yscrollcommand=scrollbar_datas.set,
            exportselection=False,
            height=5,
            relief="solid",
            bd=1,
            activestyle="none",
        )
        scrollbar_datas.config(command=self.lst_datas.yview)
        scrollbar_datas.pack(side="right", fill="y")
        self.lst_datas.pack(side="left", fill="x", expand=True)
        self.lst_datas.bind("<<ListboxSelect>>", self.on_date_select)

        # Contador de Ordens e Informações
        self.lbl_qtd_ordens = tk.Label(
            left_panel,
            text="Ordens encontradas: 0",
            font=("Segoe UI", 9, "italic"),
            bg=self.bg_color,
            fg="#64748b",
            anchor="w",
        )
        self.lbl_qtd_ordens.pack(fill="x", pady=(4, 6))

        # Botão Exportar Todos
        self.btn_exportar_todos = tk.Button(
            left_panel,
            text="📦 Exportar Todos (.txt)",
            command=self.exportar_todos_em_lote,
            font=("Segoe UI", 9, "bold"),
            bg="#475569",
            fg="white",
            activebackground="#334155",
            activeforeground="white",
            relief="flat",
            pady=7,
            cursor="hand2",
            state="disabled",
        )
        self.btn_exportar_todos.pack(fill="x", pady=(2, 0))

        # Painel Direito: Prévia da Mensagem
        right_panel = tk.LabelFrame(
            main_container,
            text=" Prévia da Mensagem ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color,
            fg="#334155",
            padx=10,
            pady=10,
        )
        right_panel.pack(side="right", fill="both", expand=True)

        # Campo de Texto da Prévia com Barra de Rolagem
        preview_frame = tk.Frame(right_panel, bg=self.bg_color)
        preview_frame.pack(fill="both", expand=True, pady=(0, 10))

        txt_scrollbar = tk.Scrollbar(preview_frame, orient="vertical")
        self.txt_preview = tk.Text(
            preview_frame,
            font=("Consolas", 10),
            wrap="word",
            relief="solid",
            bd=1,
            yscrollcommand=txt_scrollbar.set,
        )
        txt_scrollbar.config(command=self.txt_preview.yview)
        txt_scrollbar.pack(side="right", fill="y")
        self.txt_preview.pack(side="left", fill="both", expand=True)

        # Barra de Botões de Ação
        btn_frame = tk.Frame(right_panel, bg=self.bg_color)
        btn_frame.pack(fill="x")

        self.btn_copiar_direto = tk.Button(
            btn_frame,
            text="📋 Copiar Mensagem",
            command=self.copiar_texto_direto,
            font=("Segoe UI", 10, "bold"),
            bg=self.accent_color,
            fg="white",
            activebackground=self.accent_hover,
            activeforeground="white",
            relief="flat",
            pady=8,
            cursor="hand2",
            state="disabled",
        )
        self.btn_copiar_direto.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.btn_salvar_txt = tk.Button(
            btn_frame,
            text="💾 Salvar em TXT",
            command=self.salvar_txt_direto,
            font=("Segoe UI", 10, "bold"),
            bg="#334155",
            fg="white",
            activebackground="#1e293b",
            activeforeground="white",
            relief="flat",
            pady=8,
            cursor="hand2",
            state="disabled",
        )
        self.btn_salvar_txt.pack(side="left", fill="x", expand=True, padx=(5, 5))

        self.btn_mais_opcoes = tk.Button(
            btn_frame,
            text="⚡ Mais Opções...",
            command=self.abrir_janela_opcoes,
            font=("Segoe UI", 10, "bold"),
            bg="#0284c7",
            fg="white",
            activebackground="#0369a1",
            activeforeground="white",
            relief="flat",
            pady=8,
            cursor="hand2",
            state="disabled",
        )
        self.btn_mais_opcoes.pack(side="right", fill="x", expand=False, padx=(5, 0))

        # Barra de Rodapé / Status
        self.status_bar = tk.Label(
            self.root,
            text="Pronto para importar planilha.",
            font=("Segoe UI", 9),
            bg="#e2e8f0",
            fg="#475569",
            anchor="w",
            padx=12,
            pady=4,
        )
        self.status_bar.pack(side="bottom", fill="x")

    def _on_search_focus_in(self, event: Any = None) -> None:
        if self.txt_busca.get() == "🔍 Buscar fiscal...":
            self.txt_busca.delete(0, tk.END)
            self.txt_busca.config(fg="#0f172a")

    def _on_search_focus_out(self, event: Any = None) -> None:
        if not self.txt_busca.get().strip():
            self.txt_busca.delete(0, tk.END)
            self.txt_busca.insert(0, "🔍 Buscar fiscal...")
            self.txt_busca.config(fg="#64748b")

    def _clear_search(self) -> None:
        self.txt_busca.delete(0, tk.END)
        self._on_search_focus_out()
        self._on_search_change()

    def _on_search_change(self, event: Any = None) -> None:
        termo = self.txt_busca.get().strip()
        if termo == "🔍 Buscar fiscal...":
            termo = ""

        if not termo:
            self.fiscais_filtrados = list(self.fiscais_list)
        else:
            termo_lower = termo.lower()
            self.fiscais_filtrados = [
                f for f in self.fiscais_list if termo_lower in f.lower()
            ]

        self.lst_fiscais.delete(0, tk.END)
        for f in self.fiscais_filtrados:
            self.lst_fiscais.insert(tk.END, f)

        if self.fiscais_filtrados:
            self.lst_fiscais.selection_set(0)
            self.on_fiscal_select(None)
        else:
            self.lst_datas.delete(0, tk.END)
            self.txt_preview.delete("1.0", tk.END)
            self.lbl_qtd_ordens.config(text="Nenhum fiscal encontrado")
            self._set_action_buttons_state("disabled")

    def _set_action_buttons_state(
        self, state: Literal["normal", "active", "disabled"]
    ) -> None:
        self.btn_copiar_direto.config(state=state)
        self.btn_salvar_txt.config(state=state)
        self.btn_mais_opcoes.config(state=state)

    def _selecionar_aba(self, sheet_names: List[str]) -> Optional[str]:
        """Identifica a aba correta ou pergunta ao usuário se houver dúvida."""
        # 1. Busca exata por "GERAL" (insensível a maiúsculas/espaços)
        for s in sheet_names:
            if s.strip().upper() == "GERAL":
                return s

        # 2. Busca por abas contendo "GERAL"
        for s in sheet_names:
            if "GERAL" in s.strip().upper():
                return s

        # 3. Se houver apenas 1 aba no arquivo
        if len(sheet_names) == 1:
            return sheet_names[0]

        # 4. Caso haja múltiplas abas, abre janela modal para escolha
        janela_aba = tk.Toplevel(self.root)
        janela_aba.title("Selecionar Aba")
        janela_aba.geometry("360x180")
        janela_aba.resizable(False, False)
        janela_aba.transient(self.root)
        janela_aba.grab_set()

        aba_escolhida = tk.StringVar(value=sheet_names[0])

        tk.Label(
            janela_aba,
            text="A aba 'GERAL' não foi localizada automaticamente.\nEscolha qual aba deseja utilizar:",
            font=("Segoe UI", 10),
            justify="center",
            pady=12,
        ).pack()

        cb_abas = ttk.Combobox(
            janela_aba,
            textvariable=aba_escolhida,
            values=sheet_names,
            state="readonly",
            font=("Segoe UI", 10),
            width=30,
        )
        cb_abas.pack(pady=5)

        confirmado = tk.BooleanVar(value=False)

        def confirmar() -> None:
            confirmado.set(True)
            janela_aba.destroy()

        tk.Button(
            janela_aba,
            text="Confirmar",
            command=confirmar,
            font=("Segoe UI", 10, "bold"),
            bg=self.accent_color,
            fg="white",
            relief="flat",
            padx=15,
            pady=5,
            cursor="hand2",
        ).pack(pady=12)

        self.root.wait_window(janela_aba)
        return aba_escolhida.get() if confirmado.get() else None

    def importar_planilha(self) -> None:
        caminho_arquivo: str = filedialog.askopenfilename(
            title="Selecione a planilha Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls *.xlsm")],
        )
        if not caminho_arquivo:
            return

        self.root.config(cursor="wait")
        self.status_bar.config(text="Lendo planilha Excel, aguarde...")
        self.root.update_idletasks()

        try:
            xl = pd.ExcelFile(caminho_arquivo)
            nomes_abas = [str(s) for s in xl.sheet_names]
            aba = self._selecionar_aba(nomes_abas)
            if not aba:
                self.root.config(cursor="")
                self.status_bar.config(text="Importação cancelada.")
                return

            df_temp: pd.DataFrame = pd.read_excel(caminho_arquivo, sheet_name=aba)

            # Normalização e mapeamento flexível das colunas
            norm_map = {
                normalize_header_text(col): col for col in df_temp.columns
            }

            mapeamento_colunas = {
                "FISCAL": [
                    "FISCAL",
                    "NOME FISCAL",
                    "RESPONSAVEL",
                    "FISCAL RESPONSAVEL",
                ],
                "Nº ORDEM": [
                    "N ORDEM",
                    "NO ORDEM",
                    "ORDEM",
                    "NC",
                    "NUMERO DA ORDEM",
                    "NUMERO ORDEM",
                    "OS",
                    "N ORDEM DE SERVICO",
                ],
                "Nº CLIENTE": [
                    "N CLIENTE",
                    "NO CLIENTE",
                    "NUMERO DO CLIENTE",
                    "NUMERO CLIENTE",
                    "CLIENTE",
                    "UC",
                    "CONTA CONTRATO",
                ],
                "NOME DO CLIENTE": [
                    "NOME DO CLIENTE",
                    "NOME CLIENTE",
                    "CLIENTE NOME",
                    "NOME",
                    "RAZAO SOCIAL",
                ],
                "ENDEREÇO": ["ENDERECO", "LOGRADOURO", "RUA"],
                "MEDIDOR": [
                    "MEDIDOR",
                    "NO MEDIDOR",
                    "N MEDIDOR",
                    "EQUIPAMENTO",
                    "NUMERO MEDIDOR",
                    "MEDIDOR ATUAL",
                ],
                "DATA": [
                    "DATA",
                    "DATA DO RELATORIO",
                    "DATA RELAT",
                    "DATA ENVIO",
                    "DATA DIA",
                ],
                "DATA SOLIC.": [
                    "DATA SOLIC",
                    "DATA SOLICITACAO",
                    "DATA SOLIC",
                    "DATA ENTRADA",
                ],
                "COORDENADAS": [
                    "COORDENADAS",
                    "COORDENADA",
                    "LINK",
                    "MAPS",
                    "LOCALIZACAO",
                    "GPS",
                    "LINK1",
                ],
            }

            colunas_resolvidas: Dict[str, str] = {}
            for nome_padrao, variacoes in mapeamento_colunas.items():
                for var in variacoes:
                    var_norm = normalize_header_text(var)
                    if var_norm in norm_map:
                        colunas_resolvidas[nome_padrao] = norm_map[var_norm]
                        break

            # Validação: apenas FISCAL é estritamente obrigatória
            if "FISCAL" not in colunas_resolvidas:
                colunas_disponiveis = "\n".join(f"- {c}" for c in df_temp.columns)
                messagebox.showerror(
                    "Coluna Não Encontrada",
                    "A coluna de identificação do fiscal (ex: 'FISCAL') não foi encontrada na aba selecionada.\n\n"
                    f"Colunas disponíveis na planilha:\n{colunas_disponiveis}",
                )
                self.root.config(cursor="")
                self.status_bar.config(text="Erro: Coluna FISCAL ausente.")
                return

            # Busca inteligente de coluna com links do Google Maps caso COORDENADAS não tenha sido detectada
            col_coord_atual = colunas_resolvidas.get("COORDENADAS")
            tem_links = False
            if col_coord_atual is not None:
                tem_links = (
                    df_temp[col_coord_atual]
                    .astype(str)
                    .str.contains(r"https?://", regex=True, na=False)
                    .any()
                )

            if not tem_links:
                for col in df_temp.columns:
                    col_str = str(col).upper()
                    if "TEXTO" in col_str or "RETORNO" in col_str:
                        continue
                    if (
                        df_temp[col]
                        .astype(str)
                        .str.contains(r"https?://", regex=True, na=False)
                        .sum()
                        > 5
                    ):
                        colunas_resolvidas["COORDENADAS"] = col
                        break

            # Cria coluna interna de data padronizada (DD/MM/AAAA)
            col_data_ref = colunas_resolvidas.get("DATA") or colunas_resolvidas.get("DATA SOLIC.")
            if col_data_ref:
                df_temp["_DATA_CLEAN"] = df_temp[col_data_ref].apply(format_date_str)
            else:
                df_temp["_DATA_CLEAN"] = ""

            # Cria coluna interna limpa e padronizada para filtragem rápida de fiscais
            col_fiscal_orig = colunas_resolvidas["FISCAL"]
            df_temp["_FISCAL_CLEAN"] = (
                df_temp[col_fiscal_orig]
                .astype(str)
                .str.strip()
                .str.upper()
            )

            # Lista de fiscais válidos (ignora vazios e nulos)
            valores_invalidos = {"", "NAN", "NONE", "NULL", "<NA>", "0", ".", "-"}
            raw_fiscais = df_temp["_FISCAL_CLEAN"].dropna().unique()
            self.fiscais_list = sorted(
                [str(f) for f in raw_fiscais if str(f) not in valores_invalidos]
            )

            self.df = df_temp
            self.col_map = colunas_resolvidas
            self.caminho_arquivo = caminho_arquivo
            self.aba_ativa = aba

            # Atualiza lista na interface
            self.fiscais_filtrados = list(self.fiscais_list)
            self.lst_fiscais.delete(0, tk.END)
            for fiscal in self.fiscais_filtrados:
                self.lst_fiscais.insert(tk.END, fiscal)

            # Atualiza status e labels
            nome_arquivo_base = os.path.basename(caminho_arquivo)
            self.lbl_substatus.config(
                text=f"Planilha: {nome_arquivo_base} | Aba: {aba} ({len(self.df)} linhas)"
            )
            self.status_bar.config(
                text=f"Pronto. {len(self.fiscais_list)} fiscais encontrados em {len(self.df)} registros."
            )
            self.btn_exportar_todos.config(state="normal")

            if self.fiscais_list:
                self.lst_fiscais.selection_set(0)
                self.on_fiscal_select(None)

            messagebox.showinfo(
                "Sucesso",
                f"Planilha importada com sucesso!\n\n"
                f"Arquivo: {nome_arquivo_base}\n"
                f"Aba: {aba}\n"
                f"Total de registros: {len(self.df)}\n"
                f"Fiscais identificados: {len(self.fiscais_list)}",
            )

        except PermissionError:
            messagebox.showerror(
                "Arquivo Bloqueado",
                "Não foi possível abrir o arquivo.\n"
                "Ele pode estar aberto em outro programa (como o Excel). Feche-o e tente novamente.",
            )
            self.status_bar.config(text="Erro: Arquivo bloqueado por outro processo.")
        except Exception as e:
            messagebox.showerror(
                "Erro de Leitura",
                f"Ocorreu um erro ao processar a planilha:\n{str(e)}",
            )
            self.status_bar.config(text="Erro ao importar arquivo.")
        finally:
            self.root.config(cursor="")

    def on_fiscal_select(self, event: Any = None) -> None:
        selecao = self.lst_fiscais.curselection()
        if not selecao:
            return

        indice = selecao[0]
        if indice >= len(self.fiscais_filtrados):
            return

        fiscal_nome = self.fiscais_filtrados[indice]

        if self.df is None:
            return

        fiscal_upper = fiscal_nome.strip().upper()
        df_fiscal = self.df[self.df["_FISCAL_CLEAN"] == fiscal_upper]
        total_ordens = len(df_fiscal)
        self.total_ordens_fiscal_atual = total_ordens
        self.data_selecionada = None

        datas_serie = df_fiscal["_DATA_CLEAN"].dropna()
        datas_validas = [str(d) for d in datas_serie if str(d).strip() != ""]

        if datas_validas:
            contagem_datas = pd.Series(datas_validas).value_counts()
            self.datas_fiscal_map = {
                str(k): int(v) for k, v in contagem_datas.items()
            }
            self.aplicar_ordem_datas()
        else:
            self.datas_fiscal_map = {}
            self.datas_fiscal = []
            self.lst_datas.delete(0, tk.END)
            self.lst_datas.insert(tk.END, f"[Todas as datas] ({total_ordens})")
            self.lst_datas.selection_set(0)

        self.lbl_qtd_ordens.config(text=f"Ordens encontradas: {total_ordens}")
        self.status_bar.config(
            text=f"Fiscal: {fiscal_nome} | Total de ordens: {total_ordens}"
        )

        # Atualiza a prévia da mensagem
        texto: str = self.gerar_texto_mensagem(fiscal_nome, None)
        self.txt_preview.delete("1.0", tk.END)
        self.txt_preview.insert(tk.END, texto)

        self._set_action_buttons_state("normal")

    def aplicar_ordem_datas(self) -> None:
        """Ordena as datas de forma crescente ou decrescente e atualiza a interface."""
        if not self.datas_fiscal_map:
            return

        ordem = self.var_ordem_datas.get()
        reverse = ordem == "decrescente"

        datas_ordenadas = sorted(
            self.datas_fiscal_map.items(),
            key=lambda item: parse_date_key(item[0]),
            reverse=reverse,
        )
        self.datas_fiscal = datas_ordenadas

        data_anterior = self.data_selecionada

        self.lst_datas.delete(0, tk.END)
        self.lst_datas.insert(
            tk.END, f"[Todas as datas] ({self.total_ordens_fiscal_atual})"
        )

        novo_indice = 0
        for idx, (data_str, qtd) in enumerate(self.datas_fiscal, start=1):
            plural = "ordem" if qtd == 1 else "ordens"
            self.lst_datas.insert(tk.END, f"{data_str} ({qtd} {plural})")
            if data_str == data_anterior:
                novo_indice = idx

        self.lst_datas.selection_set(novo_indice)
        self.lst_datas.see(novo_indice)

    def on_date_select(self, event: Any = None) -> None:
        selecao_data = self.lst_datas.curselection()
        if not selecao_data or self.df is None:
            return

        selecao_fiscal = self.lst_fiscais.curselection()
        if not selecao_fiscal:
            return

        fiscal_nome = self.fiscais_filtrados[selecao_fiscal[0]]
        idx_data = selecao_data[0]

        if idx_data == 0 or not self.datas_fiscal:
            # "[Todas as datas]" selecionado
            self.data_selecionada = None
            df_filtrado = self.df[
                self.df["_FISCAL_CLEAN"] == fiscal_nome.strip().upper()
            ]
            qtd = len(df_filtrado)
            self.lbl_qtd_ordens.config(text=f"Ordens encontradas: {qtd}")
            self.status_bar.config(
                text=f"Fiscal: {fiscal_nome} | Todas as datas ({qtd} ordens)"
            )
        else:
            data_escolhida = self.datas_fiscal[idx_data - 1][0]
            self.data_selecionada = data_escolhida
            df_filtrado = self.df[
                (self.df["_FISCAL_CLEAN"] == fiscal_nome.strip().upper())
                & (self.df["_DATA_CLEAN"] == data_escolhida)
            ]
            qtd = len(df_filtrado)
            self.lbl_qtd_ordens.config(text=f"Ordens em {data_escolhida}: {qtd}")
            self.status_bar.config(
                text=f"Fiscal: {fiscal_nome} | Data: {data_escolhida} ({qtd} ordens)"
            )

        texto: str = self.gerar_texto_mensagem(fiscal_nome, self.data_selecionada)
        self.txt_preview.delete("1.0", tk.END)
        self.txt_preview.insert(tk.END, texto)

    def gerar_texto_mensagem(
        self, fiscal_nome: str, data_filtro: Optional[str] = None
    ) -> str:
        if self.df is None or not fiscal_nome:
            return ""

        fiscal_upper = fiscal_nome.strip().upper()
        df_fiscal = self.df[self.df["_FISCAL_CLEAN"] == fiscal_upper]

        if data_filtro:
            df_fiscal = df_fiscal[df_fiscal["_DATA_CLEAN"] == data_filtro]

        if df_fiscal.empty:
            if data_filtro:
                return f"Nenhuma ordem encontrada para o fiscal '{fiscal_nome}' na data {data_filtro}."
            return f"Nenhuma ordem encontrada para o fiscal: {fiscal_nome}"

        col_ordem = self.col_map.get("Nº ORDEM")
        col_cliente = self.col_map.get("Nº CLIENTE")
        col_nome = self.col_map.get("NOME DO CLIENTE")
        col_endereco = self.col_map.get("ENDEREÇO")
        col_medidor = self.col_map.get("MEDIDOR")
        col_coord = self.col_map.get("COORDENADAS")

        saudacao = get_time_greeting()
        linhas_mensagem = [
            f"{saudacao}, {fiscal_nome}",
        ]

        if data_filtro:
            linhas_mensagem.append(f"Data: {data_filtro}")

        linhas_mensagem.extend(
            [
                "Segue a lista de informações para verificar se está ligada/ativa, habitada, se é uma residência ou comércio e informar as leituras, nc e nf do medidor.",
                "",
                "listas abaixo:",
                "",
            ]
        )

        contador_ordens = 0
        for _, row in df_fiscal.iterrows():
            uc = format_cell_value(row.get(col_cliente)) if col_cliente else "-"
            nc = format_cell_value(row.get(col_ordem)) if col_ordem else "-"
            nome = format_cell_value(row.get(col_nome)) if col_nome else "-"
            endereco = format_cell_value(row.get(col_endereco)) if col_endereco else "-"
            medidor = format_cell_value(row.get(col_medidor)) if col_medidor else "-"

            # Ignora linhas totalmente vazias no Excel
            if uc == "-" and nc == "-" and nome == "-" and endereco == "-":
                continue

            contador_ordens += 1
            num_formatado = f"{contador_ordens:06d}"

            bloco_ordem = [
                num_formatado,
                f"uc: {uc}",
                f"nc: {nc}",
                f"nome cliente: {nome}",
                f"endereço: {endereco}",
                f"leitura atual: {medidor}",
            ]

            # Coordenadas ou link do Google Maps se disponível
            link_mapa = ""
            if col_coord:
                link_mapa = extract_maps_link(row.get(col_coord))

            if not link_mapa:
                for val in row.values:
                    link_mapa = extract_maps_link(val)
                    if link_mapa:
                        break

            if link_mapa:
                bloco_ordem.append(f"coordenadas: {link_mapa}")

            linhas_mensagem.append("\n".join(bloco_ordem))
            linhas_mensagem.append("")

        if contador_ordens == 0:
            return f"Nenhuma ordem válida encontrada para o fiscal: {fiscal_nome}"

        return "\n".join(linhas_mensagem).strip()

    def _copiar_para_clipboard(self, texto: str) -> bool:
        """Copia texto usando o Tkinter com fallback para pyperclip."""
        copiado = False
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(texto)
            self.root.update()
            copiado = True
        except Exception:
            pass

        if not copiado and pyperclip is not None:
            try:
                pyperclip.copy(texto)
                copiado = True
            except Exception:
                pass

        return copiado

    def copiar_texto_direto(self) -> None:
        texto = self.txt_preview.get("1.0", tk.END).strip()
        if not texto:
            messagebox.showwarning("Aviso", "Nenhum texto gerado para copiar.")
            return

        sucesso = self._copiar_para_clipboard(texto)
        if sucesso:
            self.status_bar.config(
                text="✔ Mensagem copiada para a Área de Transferência com sucesso!"
            )
            messagebox.showinfo(
                "Copiado",
                "Mensagem copiada para a Área de Transferência!",
            )
        else:
            messagebox.showerror(
                "Erro",
                "Não foi possível acessar a Área de Transferência do sistema.",
            )

    def salvar_txt_direto(self) -> None:
        texto = self.txt_preview.get("1.0", tk.END).strip()
        if not texto:
            messagebox.showwarning("Aviso", "Nenhum texto gerado para salvar.")
            return

        selecao = self.lst_fiscais.curselection()
        fiscal_nome = (
            self.fiscais_filtrados[selecao[0]] if selecao else "fiscal"
        )
        fiscal_slug = fiscal_nome.lower().replace(" ", "_")

        if self.data_selecionada:
            data_slug = self.data_selecionada.replace("/", "-")
            nome_arquivo_padrao = f"mensagem_{fiscal_slug}_{data_slug}.txt"
        else:
            nome_arquivo_padrao = f"mensagem_{fiscal_slug}.txt"

        caminho_salvar: str = filedialog.asksaveasfilename(
            title="Salvar como arquivo de texto",
            defaultextension=".txt",
            initialfile=nome_arquivo_padrao,
            filetypes=[("Arquivo de Texto", "*.txt")],
        )
        if caminho_salvar:
            try:
                with open(caminho_salvar, "w", encoding="utf-8") as f:
                    f.write(texto)
                self.status_bar.config(
                    text=f"Arquivo salvo: {os.path.basename(caminho_salvar)}"
                )
                messagebox.showinfo(
                    "Sucesso", f"Arquivo salvo com sucesso em:\n{caminho_salvar}"
                )
            except Exception as e:
                messagebox.showerror(
                    "Erro ao Salvar", f"Não foi possível salvar o arquivo:\n{str(e)}"
                )

    def exportar_todos_em_lote(self) -> None:
        if not self.fiscais_list or self.df is None:
            messagebox.showwarning("Aviso", "Nenhum dado importado para exportar.")
            return

        pasta_destino = filedialog.askdirectory(
            title="Selecione a pasta onde deseja salvar os arquivos de todos os fiscais"
        )
        if not pasta_destino:
            return

        self.root.config(cursor="wait")
        self.status_bar.config(text="Exportando relatórios de todos os fiscais...")
        self.root.update_idletasks()

        arquivos_gerados = 0
        erros: List[str] = []

        data_filtro = self.data_selecionada

        for fiscal in self.fiscais_list:
            try:
                texto = self.gerar_texto_mensagem(fiscal, data_filtro)
                if not texto or texto.startswith("Nenhuma ordem"):
                    continue

                nome_sanitizado = (
                    re.sub(r'[\\/*?:"<>|]', "", fiscal).strip().replace(" ", "_").lower()
                )
                if data_filtro:
                    data_slug = data_filtro.replace("/", "-")
                    nome_arquivo = f"mensagem_{nome_sanitizado}_{data_slug}.txt"
                else:
                    nome_arquivo = f"mensagem_{nome_sanitizado}.txt"

                caminho_completo = os.path.join(pasta_destino, nome_arquivo)

                with open(caminho_completo, "w", encoding="utf-8") as f:
                    f.write(texto)
                arquivos_gerados += 1
            except Exception as e:
                erros.append(f"{fiscal}: {str(e)}")

        self.root.config(cursor="")
        filtro_info = f" da data {data_filtro}" if data_filtro else ""
        self.status_bar.config(
            text=f"Exportação em lote concluída: {arquivos_gerados} arquivos gerados{filtro_info}."
        )

        if erros:
            messagebox.showwarning(
                "Exportação Concluída com Alertas",
                f"{arquivos_gerados} arquivos foram gerados com sucesso{filtro_info}.\n"
                f"Ocorreram erros nos seguintes fiscais:\n" + "\n".join(erros[:5]),
            )
        else:
            messagebox.showinfo(
                "Exportação Concluída",
                f"Todos os relatórios{filtro_info} foram exportados com sucesso!\n\n"
                f"Total de arquivos gerados: {arquivos_gerados}\n"
                f"Pasta: {pasta_destino}",
            )

    def abrir_janela_opcoes(self) -> None:
        texto: str = self.txt_preview.get("1.0", tk.END).strip()
        if not texto:
            messagebox.showwarning("Aviso", "Nenhum texto gerado para exportar.")
            return

        selecao = self.lst_fiscais.curselection()
        fiscal_nome: str = (
            self.fiscais_filtrados[selecao[0]] if selecao else "fiscal"
        )

        top_opcoes = tk.Toplevel(self.root)
        top_opcoes.title("Opções de Exportação")
        top_opcoes.geometry("380x220")
        top_opcoes.resizable(False, False)
        top_opcoes.transient(self.root)
        top_opcoes.grab_set()

        subtitulo = f"Opções para: {fiscal_nome}"
        if self.data_selecionada:
            subtitulo += f" ({self.data_selecionada})"

        tk.Label(
            top_opcoes,
            text=subtitulo,
            font=("Segoe UI", 11, "bold"),
            fg="#1e293b",
            pady=10,
        ).pack()

        def copiar_e_fechar() -> None:
            if self._copiar_para_clipboard(texto):
                messagebox.showinfo(
                    "Sucesso",
                    f"Mensagem do fiscal '{fiscal_nome}' copiada para a Área de Transferência!",
                    parent=top_opcoes,
                )
                top_opcoes.destroy()
            else:
                messagebox.showerror(
                    "Erro",
                    "Falha ao copiar para Área de Transferência.",
                    parent=top_opcoes,
                )

        def salvar_e_fechar() -> None:
            top_opcoes.destroy()
            self.salvar_txt_direto()

        btn_clipboard = tk.Button(
            top_opcoes,
            text="📋 Para área de transferência",
            command=copiar_e_fechar,
            font=("Segoe UI", 10, "bold"),
            bg=self.accent_color,
            fg="white",
            relief="flat",
            pady=8,
            padx=10,
            cursor="hand2",
        )
        btn_clipboard.pack(fill="x", padx=30, pady=5)

        btn_txt = tk.Button(
            top_opcoes,
            text="📄 Para arquivo em texto (.txt)",
            command=salvar_e_fechar,
            font=("Segoe UI", 10, "bold"),
            bg="#475569",
            fg="white",
            relief="flat",
            pady=8,
            padx=10,
            cursor="hand2",
        )
        btn_txt.pack(fill="x", padx=30, pady=5)


if __name__ == "__main__":
    root = tk.Tk()
    app = AppFiscal(root)
    root.mainloop()
