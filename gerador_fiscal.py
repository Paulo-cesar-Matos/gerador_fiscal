# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingTypeStubs=false, reportAttributeAccessIssue=false
import os
import re
import unicodedata
from datetime import datetime
from typing import Any, Dict, List, Literal, Optional, Tuple
import numpy as np  # type: ignore
import pandas as pd  # pyright: ignore[reportMissingModuleSource]
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pyperclip  # pyright: ignore[reportMissingModuleSource]
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
            if isinstance(
                dt, (datetime, pd.Timestamp)
            ):  # pyright: ignore[reportUnnecessaryIsInstance]
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
        self.root.geometry("1080x740")
        self.root.minsize(980, 640)

        self.df: Optional[pd.DataFrame] = None
        self.df_filtrado: Optional[pd.DataFrame] = None
        self.index_atual: int = 0

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

        # Paleta de Cores
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
        main_container.pack(fill="both", expand=True, padx=12, pady=(12, 5))

        # --- PAINEL ESQUERDO: Fiscais e Datas ---
        left_panel = tk.LabelFrame(
            main_container,
            text=" Mandar para: ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color,
            fg="#334155",
            padx=8,
            pady=8,
            width=270,
        )
        left_panel.pack(side="left", fill="y", padx=(0, 6))
        left_panel.pack_propagate(False)

        # Busca de Fiscais
        search_frame = tk.Frame(left_panel, bg=self.bg_color)
        search_frame.pack(fill="x", pady=(0, 4))

        self.txt_busca = tk.Entry(
            search_frame,
            font=("Segoe UI", 9),
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
            padx=4,
            cursor="hand2",
        )
        btn_clear_search.pack(side="right", padx=(2, 0))

        # Lista de Fiscais
        list_frame = tk.Frame(left_panel, bg=self.bg_color)
        list_frame.pack(fill="both", expand=True)

        sb_fiscais_y = tk.Scrollbar(list_frame, orient="vertical")
        self.lst_fiscais = tk.Listbox(
            list_frame,
            font=("Segoe UI", 9),
            selectmode=tk.SINGLE,
            yscrollcommand=sb_fiscais_y.set,
            exportselection=False,
            height=9,
            relief="solid",
            bd=1,
            activestyle="none",
        )
        sb_fiscais_y.config(command=self.lst_fiscais.yview)
        sb_fiscais_y.pack(side="right", fill="y")
        self.lst_fiscais.pack(side="left", fill="both", expand=True)
        self.lst_fiscais.bind("<<ListboxSelect>>", self.on_fiscal_select)

        # Seção de Datas
        date_header_frame = tk.Frame(left_panel, bg=self.bg_color)
        date_header_frame.pack(fill="x", pady=(8, 2))

        lbl_datas = tk.Label(
            date_header_frame,
            text="📅 Data:",
            font=("Segoe UI", 8, "bold"),
            bg=self.bg_color,
            fg="#1e293b",
        )
        lbl_datas.pack(side="left")

        rb_dec = tk.Radiobutton(
            date_header_frame,
            text="↓ Decres.",
            variable=self.var_ordem_datas,
            value="decrescente",
            command=self.aplicar_ordem_datas,
            font=("Segoe UI", 8),
            bg=str(self.bg_color),
            activebackground=self.bg_color,
            cursor="hand2",
        )
        rb_dec.pack(side="right", padx=(1, 0))

        rb_cres = tk.Radiobutton(
            date_header_frame,
            text="↑ Cres.",
            variable=self.var_ordem_datas,
            value="crescente",
            command=self.aplicar_ordem_datas,
            font=("Segoe UI", 8),
            bg=str(self.bg_color),
            activebackground=str(self.bg_color),
            cursor="hand2",
        )
        rb_cres.pack(side="right", padx=(1, 0))

        date_frame = tk.Frame(left_panel, bg=self.bg_color)
        date_frame.pack(fill="x", pady=(0, 4))

        sb_datas = tk.Scrollbar(date_frame, orient="vertical")
        self.lst_datas = tk.Listbox(
            date_frame,
            font=("Segoe UI", 8),
            selectmode=tk.SINGLE,
            yscrollcommand=sb_datas.set,
            exportselection=False,
            height=5,
            relief="solid",
            bd=1,
            activestyle="none",
        )
        sb_datas.config(command=self.lst_datas.yview)
        sb_datas.pack(side="right", fill="y")
        self.lst_datas.pack(side="left", fill="x", expand=True)
        self.lst_datas.bind("<<ListboxSelect>>", self.on_date_select)

        self.lbl_qtd_ordens = tk.Label(
            left_panel,
            text="Ordens encontradas: 0",
            font=("Segoe UI", 8, "italic"),
            bg=self.bg_color,
            fg="#64748b",
            anchor="w",
        )
        self.lbl_qtd_ordens.pack(fill="x", pady=(2, 4))

        self.btn_exportar_todos = tk.Button(
            left_panel,
            text="📦 Exportar Todos (.txt)",
            command=self.exportar_todos_em_lote,
            font=("Segoe UI", 8, "bold"),
            bg="#475569",
            fg="white",
            activebackground="#334155",
            activeforeground="white",
            relief="flat",
            pady=6,
            cursor="hand2",
            state="disabled",
        )
        self.btn_exportar_todos.pack(fill="x")

        # --- PAINEL CENTRAL: Filtros Individuais de Cliente e Endereço ---
        middle_panel = tk.LabelFrame(
            main_container,
            text=" Filtros de Busca ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color,
            fg="#334155",
            padx=8,
            pady=8,
            width=240,
        )
        middle_panel.pack(side="left", fill="y", padx=(0, 6))
        middle_panel.pack_propagate(False)

        # Filtro de Cliente
        tk.Label(
            middle_panel,
            text="🔍 Buscar Cliente:",
            font=("Segoe UI", 8, "bold"),
            bg=self.bg_color,
            fg="#1e293b",
        ).pack(anchor="w", pady=(0, 2))

        self.ent_busca_cliente = tk.Entry(
            middle_panel,
            font=("Segoe UI", 9),
            relief="solid",
            bd=1,
        )
        self.ent_busca_cliente.pack(fill="x", pady=(0, 12))
        self.ent_busca_cliente.bind("<KeyRelease>", self.aplicar_filtros)

        # Filtro de Endereço
        tk.Label(
            middle_panel,
            text="📍 Buscar Endereço:",
            font=("Segoe UI", 8, "bold"),
            bg=self.bg_color,
            fg="#1e293b",
        ).pack(anchor="w", pady=(0, 2))

        self.ent_busca_endereco = tk.Entry(
            middle_panel,
            font=("Segoe UI", 9),
            relief="solid",
            bd=1,
        )
        self.ent_busca_endereco.pack(fill="x", pady=(0, 12))
        self.ent_busca_endereco.bind("<KeyRelease>", self.aplicar_filtros)

        # Filtro de Medidor
        tk.Label(
            middle_panel,
            text="⚡ Buscar Medidor:",
            font=("Segoe UI", 8, "bold"),
            bg=self.bg_color,
            fg="#1e293b",
        ).pack(anchor="w", pady=(0, 2))

        self.ent_busca_medidor = tk.Entry(
            middle_panel,
            font=("Segoe UI", 9),
            relief="solid",
            bd=1,
        )
        self.ent_busca_medidor.pack(fill="x", pady=(0, 12))
        self.ent_busca_medidor.bind("<KeyRelease>", self.aplicar_filtros)

        # Dica explicativa
        lbl_dica = tk.Label(
            middle_panel,
            text="💡 Dica:\nAo digitar o nome do cliente ou endereço, o termo correspondente será grifado na prévia da mensagem.",
            font=("Segoe UI", 8, "italic"),
            bg=self.bg_color,
            fg="#64748b",
            justify="left",
            wraplength=210,
        )
        lbl_dica.pack(anchor="w", pady=(10, 0))

        # --- PAINEL DIREITO: Prévia e Ações ---
        right_panel = tk.LabelFrame(
            main_container,
            text=" Prévia da Mensagem ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color,
            fg="#334155",
            padx=8,
            pady=8,
        )
        right_panel.pack(side="right", fill="both", expand=True)

        preview_frame = tk.Frame(right_panel, bg=self.bg_color)
        preview_frame.pack(fill="both", expand=True, pady=(0, 6))

        txt_scrollbar = tk.Scrollbar(preview_frame, orient="vertical")
        self.txt_preview = tk.Text(
            preview_frame,
            font=("Consolas", 9),
            wrap="word",
            relief="solid",
            bd=1,
            yscrollcommand=txt_scrollbar.set,
        )
        txt_scrollbar.config(command=self.txt_preview.yview)
        txt_scrollbar.pack(side="right", fill="y")
        self.txt_preview.pack(side="left", fill="both", expand=True)

        # Estilo de realce/grifado
        self.txt_preview.tag_config(
            "highlight",
            background="#fef08a",
            foreground="#000000",
            font=("Consolas", 9, "bold"),
        )

        # Barra de Navegação (Anterior / Próximo)
        nav_frame = tk.Frame(right_panel, bg=self.bg_color)
        nav_frame.pack(fill="x", pady=(0, 6))

        self.btn_ant = tk.Button(
            nav_frame,
            text="◄ Anterior",
            command=self.reg_anterior,
            font=("Segoe UI", 8, "bold"),
            bg="#64748b",
            fg="white",
            relief="flat",
            padx=10,
            state="disabled",
            cursor="hand2",
        )
        self.btn_ant.pack(side="left")

        self.lbl_nav_pos = tk.Label(
            nav_frame,
            text="0 / 0",
            font=("Segoe UI", 9, "bold"),
            bg=self.bg_color,
            fg="#0f172a",
        )
        self.lbl_nav_pos.pack(side="left", expand=True)

        self.btn_prox = tk.Button(
            nav_frame,
            text="Próximo ►",
            command=self.proximo_reg,
            font=("Segoe UI", 8, "bold"),
            bg="#64748b",
            fg="white",
            relief="flat",
            padx=10,
            state="disabled",
            cursor="hand2",
        )
        self.btn_prox.pack(side="right")

        # Botões Principais de Ação
        btn_frame = tk.Frame(right_panel, bg=self.bg_color)
        btn_frame.pack(fill="x")

        self.btn_copiar_direto = tk.Button(
            btn_frame,
            text="📋 Copiar Mensagem",
            command=self.copiar_texto_direto,
            font=("Segoe UI", 9, "bold"),
            bg=self.accent_color,
            fg="white",
            activebackground=self.accent_hover,
            activeforeground="white",
            relief="flat",
            pady=7,
            cursor="hand2",
            state="disabled",
        )
        self.btn_copiar_direto.pack(side="left", fill="x", expand=True, padx=(0, 4))

        self.btn_salvar_txt = tk.Button(
            btn_frame,
            text="💾 Salvar em TXT",
            command=self.salvar_txt_direto,
            font=("Segoe UI", 9, "bold"),
            bg="#334155",
            fg="white",
            activebackground="#1e293b",
            activeforeground="white",
            relief="flat",
            pady=7,
            cursor="hand2",
            state="disabled",
        )
        self.btn_salvar_txt.pack(side="left", fill="x", expand=True, padx=2)

        self.btn_mais_opcoes = tk.Button(
            btn_frame,
            text="⚡ Mais Opções...",
            command=self.abrir_janela_opcoes,
            font=("Segoe UI", 9, "bold"),
            bg="#0284c7",
            fg="white",
            activebackground="#0369a1",
            activeforeground="white",
            relief="flat",
            pady=7,
            cursor="hand2",
            state="disabled",
        )
        self.btn_mais_opcoes.pack(side="right", fill="x", expand=False, padx=(4, 0))

        # Barra de Status
        self.status_bar = tk.Label(
            self.root,
            text="Pronto para importar planilha.",
            font=("Segoe UI", 8),
            bg="#e2e8f0",
            fg="#475569",
            anchor="w",
            padx=12,
            pady=3,
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
        for s in sheet_names:
            if s.strip().upper() == "GERAL":
                return s

        for s in sheet_names:
            if "GERAL" in s.strip().upper():
                return s

        if len(sheet_names) == 1:
            return sheet_names[0]

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

            norm_map = {normalize_header_text(col): col for col in df_temp.columns}

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
                ],
                "Nº CLIENTE": [
                    "N CLIENTE",
                    "NO CLIENTE",
                    "NUMERO DO CLIENTE",
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
                ],
                "DATA": ["DATA", "DATA DO RELATORIO", "DATA RELAT", "DATA ENVIO"],
                "DATA SOLIC.": ["DATA SOLIC", "DATA SOLICITACAO", "DATA ENTRADA"],
                "COORDENADAS": [
                    "COORDENADAS",
                    "COORDENADA",
                    "LINK",
                    "MAPS",
                    "LOCALIZACAO",
                    "GPS",
                ],
            }

            colunas_resolvidas: Dict[str, str] = {}
            for nome_padrao, variacoes in mapeamento_colunas.items():
                for var in variacoes:
                    var_norm = normalize_header_text(var)
                    if var_norm in norm_map:
                        colunas_resolvidas[nome_padrao] = norm_map[var_norm]
                        break

            if "FISCAL" not in colunas_resolvidas:
                colunas_disponiveis = "\n".join(f"- {c}" for c in df_temp.columns)
                messagebox.showerror(
                    "Coluna Não Encontrada",
                    f"A coluna de identificação do fiscal não foi encontrada.\n\nColunas disponíveis:\n{colunas_disponiveis}",
                )
                self.root.config(cursor="")
                self.status_bar.config(text="Erro: Coluna FISCAL ausente.")
                return

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

            col_data_ref = colunas_resolvidas.get("DATA") or colunas_resolvidas.get(
                "DATA SOLIC."
            )
            if col_data_ref:
                df_temp["_DATA_CLEAN"] = df_temp[col_data_ref].apply(format_date_str)
            else:
                df_temp["_DATA_CLEAN"] = ""

            col_fiscal_orig = colunas_resolvidas["FISCAL"]
            df_temp["_FISCAL_CLEAN"] = (
                df_temp[col_fiscal_orig].astype(str).str.strip().str.upper()
            )

            valores_invalidos = {"", "NAN", "NONE", "NULL", "<NA>", "0", ".", "-"}
            raw_fiscais = df_temp["_FISCAL_CLEAN"].dropna().unique()
            self.fiscais_list = sorted(
                [str(f) for f in raw_fiscais if str(f) not in valores_invalidos]
            )

            self.df = df_temp
            self.col_map = colunas_resolvidas
            self.caminho_arquivo = caminho_arquivo
            self.aba_ativa = aba

            self.fiscais_filtrados = list(self.fiscais_list)
            self.lst_fiscais.delete(0, tk.END)
            for fiscal in self.fiscais_filtrados:
                self.lst_fiscais.insert(tk.END, fiscal)

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

        except Exception as e:
            messagebox.showerror(
                "Erro de Leitura", f"Erro ao processar planilha:\n{str(e)}"
            )
            self.status_bar.config(text="Erro ao importar arquivo.")
        finally:
            self.root.config(cursor="")

    def on_fiscal_select(self, event: Any = None) -> None:
        if self.df is None:
            return

        # Se foi um clique manual do usuário na lista, limpa a busca global
        if event is not None:
            if hasattr(self, "ent_busca_cliente"):
                self.ent_busca_cliente.delete(0, tk.END)
            if hasattr(self, "ent_busca_endereco"):
                self.ent_busca_endereco.delete(0, tk.END)
            if hasattr(self, "ent_busca_medidor"):
                self.ent_busca_medidor.delete(0, tk.END)

        selecao = self.lst_fiscais.curselection()
        if not selecao:
            return

        fiscal_nome = self.fiscais_filtrados[selecao[0]]
        fiscal_upper = fiscal_nome.strip().upper()
        df_fiscal = self.df[self.df["_FISCAL_CLEAN"] == fiscal_upper]

        total_ordens = len(df_fiscal)
        self.total_ordens_fiscal_atual = total_ordens

        # Reseta a data ao trocar de fiscal manualmente
        if event is not None:
            self.data_selecionada = None

        datas_serie = df_fiscal["_DATA_CLEAN"].dropna()
        datas_validas = [str(d) for d in datas_serie if str(d).strip() != ""]

        if datas_validas:
            contagem_datas = pd.Series(datas_validas).value_counts()
            self.datas_fiscal_map = {str(k): int(v) for k, v in contagem_datas.items()}
            self.aplicar_ordem_datas()
        else:
            self.datas_fiscal_map = {}
            self.datas_fiscal = []
            self.lst_datas.delete(0, tk.END)
            self.lst_datas.insert(tk.END, f"[Todas as datas] ({total_ordens})")
            self.lst_datas.selection_set(0)

        self.aplicar_filtros()

    def aplicar_ordem_datas(self) -> None:
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
        if not selecao_data:
            return

        # Se foi um clique manual na lista de datas, limpa a busca global
        if event is not None:
            if hasattr(self, "ent_busca_cliente"):
                self.ent_busca_cliente.delete(0, tk.END)
            if hasattr(self, "ent_busca_endereco"):
                self.ent_busca_endereco.delete(0, tk.END)
            if hasattr(self, "ent_busca_medidor"):
                self.ent_busca_medidor.delete(0, tk.END)

        idx_data = selecao_data[0]
        if idx_data == 0 or not self.datas_fiscal:
            self.data_selecionada = None
        else:
            self.data_selecionada = self.datas_fiscal[idx_data - 1][0]

        self.aplicar_filtros()

    def aplicar_filtros(self, event: Any = None) -> None:
        """Busca global e universal em toda a planilha."""
        if self.df is None:
            return

        busca_cli = (
            self.ent_busca_cliente.get().strip().lower()
            if hasattr(self, "ent_busca_cliente")
            else ""
        )
        busca_end = (
            self.ent_busca_endereco.get().strip().lower()
            if hasattr(self, "ent_busca_endereco")
            else ""
        )
        busca_med = (
            self.ent_busca_medidor.get().strip().lower()
            if hasattr(self, "ent_busca_medidor")
            else ""
        )

        tem_busca_global = bool(busca_cli or busca_end or busca_med)

        if tem_busca_global:
            # Inicia com uma cópia de toda a planilha
            df_f = self.df.copy()

            # 1. Filtro por Cliente (se preenchido)
            if busca_cli:
                col_nome = self.col_map.get("NOME DO CLIENTE")
                if col_nome and col_nome in df_f.columns:
                    df_f = df_f[
                        df_f[col_nome]
                        .astype(str)
                        .str.lower()
                        .str.contains(busca_cli, na=False)
                    ]

            # 2. Filtro por Endereço (se preenchido)
            if busca_end:
                col_end = self.col_map.get("ENDEREÇO")
                if col_end and col_end in df_f.columns:
                    df_f = df_f[
                        df_f[col_end]
                        .astype(str)
                        .str.lower()
                        .str.contains(busca_end, na=False)
                    ]

            # 3. Busca Universal por Medidor/Número (Varre TODAS as colunas da planilha inteira)
            if busca_med:
                # Usa regex com limites (\b) para garantir que busque o número exato, 
                # evitando que pegue pedaços de coordenadas, URLs ou outros números.
                padrao_regex = rf"\b{re.escape(busca_med)}\b"
                
                condicao_unificada = False
                for col in df_f.columns:
                    col_str = df_f[col].apply(format_cell_value).str.lower()
                    condicao_unificada |= col_str.str.contains(padrao_regex, regex=True, na=False)
                
                df_f = df_f[condicao_unificada]

            self.df_filtrado = df_f
            self.index_atual = 0

            # Sincroniza automaticamente o Fiscal e a Data do primeiro resultado encontrado
            if not df_f.empty:
                primeiro_registro = df_f.iloc[0]
                fiscal_encontrado = (
                    str(primeiro_registro.get("_FISCAL_CLEAN", "")).strip().upper()
                )
                data_encontrada = str(primeiro_registro.get("_DATA_CLEAN", ""))

                # Seleciona o fiscal na lista da esquerda
                for idx, f_nome in enumerate(self.fiscais_filtrados):
                    if f_nome.strip().upper() == fiscal_encontrado:
                        self.lst_fiscais.selection_clear(0, tk.END)
                        self.lst_fiscais.selection_set(idx)
                        self.lst_fiscais.see(idx)
                        break

                # Atualiza as datas do fiscal encontrado
                df_fiscal_completo = self.df[
                    self.df["_FISCAL_CLEAN"] == fiscal_encontrado
                ]
                self.total_ordens_fiscal_atual = len(df_fiscal_completo)
                datas_serie = df_fiscal_completo["_DATA_CLEAN"].dropna()
                datas_validas = [str(d) for d in datas_serie if str(d).strip() != ""]

                if datas_validas:
                    self.datas_fiscal_map = {
                        str(k): int(v)
                        for k, v in pd.Series(datas_validas).value_counts().items()
                    }
                else:
                    self.datas_fiscal_map = {}

                self.data_selecionada = data_encontrada
                self.aplicar_ordem_datas()

        else:
            # Modo padrão (navegação manual por cliques na lista)
            selecao_fiscal = self.lst_fiscais.curselection()
            if not selecao_fiscal:
                self.df_filtrado = pd.DataFrame()
                self.index_atual = 0
                self.exibir_mensagem_individual()
                return

            fiscal_nome = self.fiscais_filtrados[selecao_fiscal[0]]
            df_f = self.df[
                self.df["_FISCAL_CLEAN"] == fiscal_nome.strip().upper()
            ].copy()

            if self.data_selecionada:
                df_f = df_f[df_f["_DATA_CLEAN"] == self.data_selecionada]

            self.df_filtrado = df_f
            self.index_atual = 0

        # Atualiza a interface e contadores
        qtd = len(self.df_filtrado) if self.df_filtrado is not None else 0
        self.lbl_qtd_ordens.config(text=f"Ordens encontradas: {qtd}")

        if tem_busca_global:
            self.status_bar.config(
                text=f"Busca Universal Ativada | Total exibido: {qtd} ordens"
            )
        else:
            fiscal_selecionado = (
                self.fiscais_filtrados[self.lst_fiscais.curselection()[0]]
                if self.lst_fiscais.curselection()
                else ""
            )
            self.status_bar.config(
                text=f"Fiscal: {fiscal_selecionado} | Total exibido: {qtd} ordens"
            )

        self.exibir_mensagem_individual()

    def exibir_mensagem_individual(self) -> None:
        """Gera e exibe a mensagem para o registro atual grifando os termos buscados."""
        if self.df_filtrado is None or self.df_filtrado.empty:
            self.txt_preview.delete("1.0", tk.END)
            self.txt_preview.insert(
                tk.END, "Nenhum registro encontrado para os filtros selecionados."
            )
            self.lbl_nav_pos.config(text="0 / 0")
            self.btn_ant.config(state="disabled")
            self.btn_prox.config(state="disabled")
            self._set_action_buttons_state("disabled")
            return

        total = len(self.df_filtrado)
        row = self.df_filtrado.iloc[self.index_atual]

        col_ordem = self.col_map.get("Nº ORDEM")
        col_cliente = self.col_map.get("Nº CLIENTE")
        col_nome = self.col_map.get("NOME DO CLIENTE")
        col_endereco = self.col_map.get("ENDEREÇO")
        col_coord = self.col_map.get("COORDENADAS")
        col_medidor = self.col_map.get("MEDIDOR")

        selecao_f = self.lst_fiscais.curselection()
        fiscal_nome = self.fiscais_filtrados[selecao_f[0]] if selecao_f else "Fiscal"

        uc = format_cell_value(row.get(col_cliente)) if col_cliente else "-"
        nc = format_cell_value(row.get(col_ordem)) if col_ordem else "-"
        medidor = format_cell_value(row.get(col_medidor)) if col_medidor else "-"
        nome = format_cell_value(row.get(col_nome)) if col_nome else "-"
        endereco = format_cell_value(row.get(col_endereco)) if col_endereco else "-"

        num_formatado = f"{self.index_atual + 1:06d}"
        saudacao = get_time_greeting()

        linhas = [
            f"{saudacao}, {fiscal_nome}",
        ]

        if self.data_selecionada:
            linhas.append(f"Data: {self.data_selecionada}")

        linhas.extend(
            [
                "Segue a lista de informações para verificar se está ligada/ativa, habitada, se é uma residência ou comércio e informar as leituras, nc e nf do medidor.",
                "",
                "listas abaixo:",
                "",
                num_formatado,
                f"uc: {uc}",
                f"nc: {nc}",
                f"medidor: {medidor}",
                f"nome cliente: {nome}",
                f"endereço: {endereco}",
            ]
        )

        link_mapa = extract_maps_link(row.get(col_coord)) if col_coord else ""
        if not link_mapa:
            for val in row.values:
                link_mapa = extract_maps_link(val)
                if link_mapa:
                    break

        if link_mapa:
            linhas.append(f"coordenadas: {link_mapa}")

        texto_mensagem = "\n".join(linhas)

        self.txt_preview.delete("1.0", tk.END)
        self.txt_preview.insert(tk.END, texto_mensagem)

        # Grifa/Realça o texto buscado
        self._grifar_termos_buscados()

        self.lbl_nav_pos.config(text=f"{self.index_atual + 1} / {total}")
        self.btn_ant.config(state="normal" if self.index_atual > 0 else "disabled")
        self.btn_prox.config(
            state="normal" if self.index_atual < total - 1 else "disabled"
        )
        self._set_action_buttons_state("normal")

    def _grifar_termos_buscados(self) -> None:
        """Aplica o fundo amarelo nas palavras correspondentes às buscas de Cliente, Endereço e Medidor/NC."""
        self.txt_preview.tag_remove("highlight", "1.0", tk.END)

        termos = [
            self.ent_busca_cliente.get().strip(),
            self.ent_busca_endereco.get().strip(),
            (
                self.ent_busca_medidor.get().strip()
                if hasattr(self, "ent_busca_medidor")
                else ""
            ),
        ]

        for termo in termos:
            if not termo or len(termo) < 2:
                continue

            pos_inicio = "1.0"
            while True:
                pos_match = self.txt_preview.search(
                    termo, pos_inicio, stopindex=tk.END, nocase=True
                )
                if not pos_match:
                    break
                pos_fim = f"{pos_match}+{len(termo)}c"
                self.txt_preview.tag_add("highlight", pos_match, pos_fim)
                pos_inicio = pos_fim

    def proximo_reg(self) -> None:
        if (
            self.df_filtrado is not None
            and self.index_atual < len(self.df_filtrado) - 1
        ):
            self.index_atual += 1
            self.exibir_mensagem_individual()

    def reg_anterior(self) -> None:
        if self.index_atual > 0:
            self.index_atual -= 1
            self.exibir_mensagem_individual()

    def _copiar_para_clipboard(self, texto: str) -> bool:
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

        if self._copiar_para_clipboard(texto):
            self.status_bar.config(text="✔ Mensagem copiada com sucesso!")
            messagebox.showinfo(
                "Copiado", "Mensagem copiada para a Área de Transferência!"
            )

    def salvar_txt_direto(self) -> None:
        texto = self.txt_preview.get("1.0", tk.END).strip()
        if not texto:
            messagebox.showwarning("Aviso", "Nenhum texto gerado para salvar.")
            return

        selecao = self.lst_fiscais.curselection()
        fiscal_nome = self.fiscais_filtrados[selecao[0]] if selecao else "fiscal"
        fiscal_slug = fiscal_nome.lower().replace(" ", "_")

        nome_arquivo = f"mensagem_{fiscal_slug}_ordem_{self.index_atual + 1}.txt"

        caminho_salvar: str = filedialog.asksaveasfilename(
            title="Salvar como arquivo de texto",
            defaultextension=".txt",
            initialfile=nome_arquivo,
            filetypes=[("Arquivo de Texto", "*.txt")],
        )
        if caminho_salvar:
            try:
                with open(caminho_salvar, "w", encoding="utf-8") as f:
                    f.write(texto)
                self.status_bar.config(
                    text=f"Arquivo salvo: {os.path.basename(caminho_salvar)}"
                )
                messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{caminho_salvar}")
            except Exception as e:
                messagebox.showerror(
                    "Erro ao Salvar", f"Não foi possível salvar:\n{str(e)}"
                )

    def exportar_todos_em_lote(self) -> None:
        if self.df_filtrado is None or self.df_filtrado.empty:
            messagebox.showwarning("Aviso", "Nenhum dado filtrado para exportar.")
            return

        caminho = filedialog.asksaveasfilename(
            title="Salvar todas as ordens filtradas",
            defaultextension=".txt",
            filetypes=[("Arquivo de Texto", "*.txt")],
            initialfile="ordens_filtradas_exportacao.txt",
        )
        if not caminho:
            return

        self.root.config(cursor="wait")
        try:
            texto_completo = ""
            total = len(self.df_filtrado)

            for i in range(total):
                self.index_atual = i
                self.exibir_mensagem_individual()
                texto_completo += (
                    self.txt_preview.get("1.0", tk.END).strip()
                    + "\n\n"
                    + ("=" * 40)
                    + "\n\n"
                )

            with open(caminho, "w", encoding="utf-8") as f:
                f.write(texto_completo)

            messagebox.showinfo(
                "Sucesso", f"Todas as {total} ordens filtradas foram exportadas!"
            )
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar lote:\n{str(e)}")
        finally:
            self.root.config(cursor="")

    def abrir_janela_opcoes(self) -> None:
        texto: str = self.txt_preview.get("1.0", tk.END).strip()
        if not texto:
            messagebox.showwarning("Aviso", "Nenhum texto para exportar.")
            return

        top_opcoes = tk.Toplevel(self.root)
        top_opcoes.title("Opções de Exportação")
        top_opcoes.geometry("380x200")
        top_opcoes.resizable(False, False)
        top_opcoes.transient(self.root)
        top_opcoes.grab_set()

        tk.Label(
            top_opcoes,
            text="Escolha onde deseja salvar/copiar:",
            font=("Segoe UI", 11, "bold"),
            fg="#1e293b",
            pady=15,
        ).pack()

        def copiar_e_fechar() -> None:
            self.copiar_texto_direto()
            top_opcoes.destroy()

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
