"""
RankingLojas.py
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Importa os dados dos Excel para um banco SQLite (lojas.db)
• Exibe interface gráfica com:
    - Tabela de ranking por loja
    - Gráfico de barras de faturamento
    - Aba de detalhes por loja (vendedores)
    - Botão para enviar ranking por e-mail
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Dependências:  pip install pandas openpyxl sqlalchemy yagmail matplotlib
"""

from logging import root
import os
import tkinter as tk
from tkinter import ttk, messagebox

import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from sqlalchemy import create_engine, text
import yagmail
from chave import senha
# ─────────────────────────────────────────────────────────────
#  CONFIGURAÇÕES  (ajuste conforme necessário)
# ─────────────────────────────────────────────────────────────
GMAIL_USER  = "mersonet@gmail.com"
GMAIL_SENHA = senha         # senha de app do Gmail em arquivo separado
EMAIL_TO    = ["emersonlopes1329@gmail.com"]
EMAIL_CC    = ["mersonet@hotmail.com"]
EMAIL_BCC   = ["lopes-ef@hotmail.com"]

LOJAS       = ["BH", "DF", "Manaus", "Rio", "Salvador", "SP"]

def _detectar_pasta():
    """Detecta automaticamente onde estão os arquivos Excel."""
    candidatas = []
    try:
        candidatas.append(os.path.dirname(os.path.abspath(__file__)))
    except NameError:
        pass
    candidatas.append(os.getcwd())
    for pasta in candidatas:
        for loja in ["BH", "SP"]:
            if (os.path.exists(os.path.join(pasta, f"Loja_{loja}.xlsx")) or
                    os.path.exists(os.path.join(pasta, f"Loja {loja}.xlsx"))):
                return pasta
    return candidatas[0] if candidatas else os.getcwd()

PASTA_EXCEL = _detectar_pasta()
DB_PATH     = os.path.join(PASTA_EXCEL, "lojas.db")

CORES = {
    "bg":        "#1e1e2e",
    "card":      "#2a2a3e",
    "accent":    "#7c3aed",
    "accent2":   "#06b6d4",
    "text":      "#e2e8f0",
    "subtext":   "#94a3b8",
    "success":   "#10b981",
    "warning":   "#f59e0b",
    "danger":    "#ef4444",
    "bar_cores": ["#7c3aed","#06b6d4","#10b981","#f59e0b","#ef4444","#ec4899"],
}


# ─────────────────────────────────────────────────────────────
#  BANCO DE DADOS
# ─────────────────────────────────────────────────────────────
engine = create_engine(f"sqlite:///{DB_PATH}")

def importar_excel_para_db():
    """Lê cada xlsx e salva/atualiza no banco SQLite."""
    importados = []
    nao_encontrados = []

    with engine.begin() as conn:
        for loja in LOJAS:
            # Tenta os dois formatos de nome de arquivo
            candidatos = [
                os.path.join(PASTA_EXCEL, f"Loja_{loja}.xlsx"),
                os.path.join(PASTA_EXCEL, f"Loja {loja}.xlsx"),
            ]
            caminho = next((c for c in candidatos if os.path.exists(c)), None)

            if caminho is None:
                nao_encontrados.append(loja)
                print(f"[AVISO] Arquivo não encontrado para loja '{loja}'.")
                print(f"        Caminhos tentados:")
                for c in candidatos:
                    print(f"          {c}")
                continue

            df = pd.read_excel(caminho).dropna(subset=["Vendas"])
            df["Loja"] = loja
            df.to_sql(f"loja_{loja.lower()}", conn,
                      if_exists="replace", index=False)
            importados.append(loja)
            print(f"  ✅ {loja}: {len(df)} registros importados ({caminho})")

    print(f"\n✅ Importação concluída: {importados}")
    if nao_encontrados:
        print(f"⚠️  Não encontrados: {nao_encontrados}")
        print(f"   Pasta procurada: {PASTA_EXCEL}")
        print(f"   Arquivos na pasta: {os.listdir(PASTA_EXCEL)}")


def get_ranking() -> pd.DataFrame:
    """Retorna DataFrame com faturamento e lucro por loja, ordenado desc."""
    rows = []
    with engine.connect() as conn:
        for loja in LOJAS:
            try:
                res = conn.execute(
                    text(f"SELECT SUM(Vendas), SUM(Lucro) FROM loja_{loja.lower()}")
                ).fetchone()
                if res and res[0] is not None:
                    rows.append({"Loja": loja,
                                 "Faturamento": res[0],
                                 "Lucro":       res[1] or 0})
            except Exception as e:
                print(f"[get_ranking] Erro na loja {loja}: {e}")

    if not rows:
        print("\n⚠️  Nenhum dado encontrado no banco.")
        print("   Verifique se os arquivos Excel estão na mesma pasta do script.")
        print(f"   Pasta atual: {PASTA_EXCEL}")
        print(f"   Arquivos encontrados: {[f for f in os.listdir(PASTA_EXCEL) if f.endswith('.xlsx')]}")
        return pd.DataFrame(columns=["Loja", "Faturamento", "Lucro"])

    df = pd.DataFrame(rows).sort_values("Faturamento", ascending=False).reset_index(drop=True)
    df.index += 1
    return df


def get_detalhes_loja(loja: str) -> pd.DataFrame:
    """Retorna vendas agrupadas por vendedor para uma loja."""
    with engine.connect() as conn:
        df = pd.read_sql(
            text(f"SELECT Vendedor, SUM(Vendas) as Vendas, SUM(Lucro) as Lucro "
                 f"FROM loja_{loja.lower()} GROUP BY Vendedor ORDER BY Vendas DESC"),
            conn
        )
    return df


# ─────────────────────────────────────────────────────────────
#  E-MAIL
# ─────────────────────────────────────────────────────────────
def gerar_html_email(ranking: pd.DataFrame) -> str:
    linhas = ""
    cores_rank = [CORES["success"], CORES["warning"], CORES["accent2"]]
    for i, row in ranking.iterrows():
        cor = cores_rank[i - 1] if i <= 3 else "#555"
        linhas += f"""
        <tr>
          <td style='padding:8px; text-align:center; color:{cor}; font-weight:bold;'>#{i}</td>
          <td style='padding:8px;'>{row['Loja']}</td>
          <td style='padding:8px; text-align:right;'>${row['Faturamento']:,.2f}</td>
          <td style='padding:8px; text-align:right;'>${row['Lucro']:,.2f}</td>
        </tr>"""

    return f"""
    <html><body style='font-family:Arial,sans-serif; background:#f4f4f4; padding:20px;'>
      <div style='max-width:600px; margin:auto; background:white; border-radius:10px;
                  box-shadow:0 2px 8px rgba(0,0,0,.15); overflow:hidden;'>
        <div style='background:#7c3aed; padding:20px; color:white; text-align:center;'>
          <h2 style='margin:0;'>📊 Ranking de Vendas</h2>
        </div>
        <div style='padding:20px;'>
          <p>Prezados,<br>Segue o ranking atualizado de faturamento das lojas:</p>
          <table style='width:100%; border-collapse:collapse;'>
            <thead>
              <tr style='background:#7c3aed; color:white;'>
                <th style='padding:8px;'>#</th>
                <th style='padding:8px;'>Loja</th>
                <th style='padding:8px;'>Faturamento</th>
                <th style='padding:8px;'>Lucro</th>
              </tr>
            </thead>
            <tbody>{linhas}</tbody>
          </table>
          <p style='margin-top:20px;'>Qualquer dúvida, estou à disposição.<br>
          <strong>Att., Lira</strong></p>
        </div>
      </div>
    </body></html>"""


def enviar_email(ranking: pd.DataFrame):
    try:
        html = gerar_html_email(ranking)
        yag = yagmail.SMTP(GMAIL_USER, GMAIL_SENHA)
        yag.send(
            to=EMAIL_TO, cc=EMAIL_CC, bcc=EMAIL_BCC,
            subject="📊 Ranking de Vendas das Lojas",
            contents=[html]
        )
        messagebox.showinfo("E-mail", "✅ E-mail enviado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao enviar e-mail:\n{e}")


# ─────────────────────────────────────────────────────────────
#  INTERFACE GRÁFICA
# ─────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🏪 Ranking de Lojas")
        self.geometry("900x640")
        self.configure(bg=CORES["bg"])
        self.resizable(True, True)

        # Configura o comportamento ao fechar a janela
        self.protocol("WM_DELETE_WINDOW", self.fechar_app)

        self.ranking = get_ranking()
        self._build_header()
        self._build_notebook()
        self._build_footer()

    def fechar_app(self):
        """Função chamada ao clicar no X da janela"""
        self.destroy()
        self.quit()


    # ── HEADER ──────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=CORES["accent"], pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📊  Ranking de Vendas por Loja",
                 font=("Segoe UI", 16, "bold"),
                 bg=CORES["accent"], fg="white").pack()
 #       tk.Label(hdr, text=f"Banco de dados: {DB_PATH}",
 #                font=("Segoe UI", 8),
 #              bg=CORES["accent"], fg="#ddd6fe").pack()

    # ── NOTEBOOK (abas) ──────────────────────────────────────
    def _build_notebook(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook",        background=CORES["bg"], borderwidth=0)
        style.configure("TNotebook.Tab",    background=CORES["card"],
                         foreground=CORES["subtext"], padding=[14, 6],
                         font=("Segoe UI", 10))
        style.map("TNotebook.Tab",
                  background=[("selected", CORES["accent"])],
                  foreground=[("selected", "white")])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=10)

        self.frm_ranking  = tk.Frame(nb, bg=CORES["bg"])
        self.frm_grafico  = tk.Frame(nb, bg=CORES["bg"])
        self.frm_detalhes = tk.Frame(nb, bg=CORES["bg"])

        nb.add(self.frm_ranking,  text="🏆  Ranking")
        nb.add(self.frm_grafico,  text="📈  Gráfico")
        nb.add(self.frm_detalhes, text="🔍  Detalhes por Loja")

        self._build_aba_ranking()
        self._build_aba_grafico()
        self._build_aba_detalhes()

    # ── ABA RANKING ──────────────────────────────────────────
    def _build_aba_ranking(self):
        frm = self.frm_ranking

        # Cards de resumo no topo
        cards = tk.Frame(frm, bg=CORES["bg"])
        cards.pack(fill="x", padx=14, pady=(12, 4))

        total_fat = self.ranking["Faturamento"].sum()
        total_luc = self.ranking["Lucro"].sum()
        melhor    = self.ranking.iloc[0]["Loja"]

        for titulo, valor in [
            ("💰 Faturamento Total",  f"${total_fat:,.0f}"),
            ("📈 Lucro Total",        f"${total_luc:,.0f}"),
            ("🥇 Melhor Loja",        melhor),
        ]:
            card = tk.Frame(cards, bg=CORES["card"], padx=18, pady=10,
                            relief="flat", bd=0)
            card.pack(side="left", expand=True, fill="both", padx=6)
            tk.Label(card, text=titulo, bg=CORES["card"],
                     fg=CORES["subtext"], font=("Segoe UI", 9)).pack()
            tk.Label(card, text=valor, bg=CORES["card"],
                     fg=CORES["accent2"], font=("Segoe UI", 14, "bold")).pack()

        # Tabela Treeview
        cols = ("#", "Loja", "Faturamento", "Lucro")
        tree_frame = tk.Frame(frm, bg=CORES["bg"])
        tree_frame.pack(fill="both", expand=True, padx=14, pady=8)

        style = ttk.Style()
        style.configure("Rank.Treeview",
                         background=CORES["card"],
                         foreground=CORES["text"],
                         fieldbackground=CORES["card"],
                         rowheight=32,
                         font=("Segoe UI", 10))
        style.configure("Rank.Treeview.Heading",
                         background=CORES["accent"],
                         foreground="white",
                         font=("Segoe UI", 10, "bold"))
        style.map("Rank.Treeview",
                  background=[("selected", CORES["accent"])])

        tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                            style="Rank.Treeview")
        for c in cols:
            tree.heading(c, text=c)
        tree.column("#",           width=50,  anchor="center")
        tree.column("Loja",        width=120, anchor="center")
        tree.column("Faturamento", width=180, anchor="e")
        tree.column("Lucro",       width=150, anchor="e")
        tree.pack(fill="both", expand=True)

        medalhas = ["🥇", "🥈", "🥉"]
        for i, row in self.ranking.iterrows():
            pos   = medalhas[i - 1] if i <= 3 else f"#{i}"
            fat   = f"${row['Faturamento']:,.2f}"
            lucro = f"${row['Lucro']:,.2f}"
            tree.insert("", "end", values=(pos, row["Loja"], fat, lucro))

    # ── ABA GRÁFICO ──────────────────────────────────────────
    def _build_aba_grafico(self):
        fig, ax = plt.subplots(figsize=(8, 4.2), facecolor=CORES["bg"])
        ax.set_facecolor(CORES["card"])

        lojas = self.ranking["Loja"].tolist()
        vals  = self.ranking["Faturamento"].tolist()
        cores = CORES["bar_cores"][:len(lojas)]

        bars = ax.barh(lojas, vals, color=cores, height=0.55, edgecolor="none")

        for bar, val in zip(bars, vals):
            ax.text(bar.get_width() + max(vals) * 0.01, bar.get_y() + bar.get_height() / 2,
                    f"${val:,.0f}", va="center", color=CORES["text"],
                    fontsize=9, fontweight="bold")

        ax.set_xlabel("Faturamento (R$)", color=CORES["subtext"])
        ax.set_title("Faturamento por Loja", color=CORES["text"],
                     fontsize=13, fontweight="bold", pad=12)
        ax.tick_params(colors=CORES["subtext"])
        ax.spines[:].set_visible(False)
        ax.xaxis.label.set_color(CORES["subtext"])
        ax.set_xlim(0, max(vals) * 1.18)

        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.frm_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)

    # ── ABA DETALHES ─────────────────────────────────────────
    def _build_aba_detalhes(self):
        frm = self.frm_detalhes

        top = tk.Frame(frm, bg=CORES["bg"])
        top.pack(fill="x", padx=14, pady=10)

        tk.Label(top, text="Selecione a loja:", bg=CORES["bg"],
                 fg=CORES["text"], font=("Segoe UI", 10)).pack(side="left")

        self.loja_var = tk.StringVar(value=LOJAS[0])
        combo = ttk.Combobox(top, textvariable=self.loja_var,
                             values=LOJAS, state="readonly", width=14)
        combo.pack(side="left", padx=8)
        combo.bind("<<ComboboxSelected>>", lambda e: self._atualizar_detalhes())

        # Treeview de vendedores
        cols = ("Vendedor", "Vendas", "Lucro")
        style = ttk.Style()
        style.configure("Det.Treeview",
                         background=CORES["card"], foreground=CORES["text"],
                         fieldbackground=CORES["card"], rowheight=28,
                         font=("Segoe UI", 10))
        style.configure("Det.Treeview.Heading",
                         background=CORES["accent2"], foreground="white",
                         font=("Segoe UI", 10, "bold"))
        style.map("Det.Treeview", background=[("selected", CORES["accent"])])

        self.tree_det = ttk.Treeview(frm, columns=cols,
                                      show="headings", style="Det.Treeview")
        for c in cols:
            self.tree_det.heading(c, text=c)
        self.tree_det.column("Vendedor", width=200, anchor="w")
        self.tree_det.column("Vendas",   width=160, anchor="e")
        self.tree_det.column("Lucro",    width=140, anchor="e")
        self.tree_det.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        self._atualizar_detalhes()

    def _atualizar_detalhes(self):
        self.tree_det.delete(*self.tree_det.get_children())
        df = get_detalhes_loja(self.loja_var.get())
        for _, row in df.iterrows():
            self.tree_det.insert("", "end", values=(
                row["Vendedor"],
                f"${row['Vendas']:,.2f}",
                f"${row['Lucro']:,.2f}"
            ))

    # ── FOOTER ───────────────────────────────────────────────
    def _build_footer(self):
        ftr = tk.Frame(self, bg=CORES["card"], pady=8)
        ftr.pack(fill="x", side="bottom")

        tk.Button(ftr, text="📧  Enviar Ranking por E-mail",
                  bg=CORES["accent"], fg="white", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=18, pady=6,
                  activebackground="#6d28d9", cursor="hand2",
                  command=lambda: enviar_email(self.ranking)
                  ).pack(side="right", padx=16)

        tk.Button(ftr, text="🔄  Reimportar Excel → BD",
                  bg=CORES["success"], fg="white", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=14, pady=6,
                  activebackground="#059669", cursor="hand2",
                  command=self._reimportar
                  ).pack(side="right", padx=6)

        tk.Label(ftr, text="lojas.db  •  SQLite + SQLAlchemy",
                 bg=CORES["card"], fg=CORES["subtext"],
                 font=("Segoe UI", 8)).pack(side="left", padx=16)

    def _reimportar(self):
        importar_excel_para_db()
        self.ranking = get_ranking()
        messagebox.showinfo("Banco de Dados",
                            "✅ Dados reimportados!\nReinicie o programa para ver as atualizações.")


# ─────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("⏳ Importando Excel para o banco de dados...")
    importar_excel_para_db()

    print("\n📊 Ranking atual:")
    ranking = get_ranking()
    ranking_print = ranking.copy()
    ranking_print["Faturamento"] = ranking_print["Faturamento"].map("${:,.2f}".format)
    ranking_print["Lucro"]       = ranking_print["Lucro"].map("${:,.2f}".format)
    print(ranking_print.to_string())

    print("\n🖥️  Abrindo interface gráfica...")
    app = App()
    app.mainloop()
    
    print("\nInterface fechada. Voltando ao terminal...")
