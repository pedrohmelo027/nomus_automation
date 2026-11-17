import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from src.leitor_cad import ler_bom_cad
from src.regras import carregar_config, processar_bom_para_planilha_trabalho
from src.planilha_trabalho import salvar_planilha_trabalho, ler_planilha_trabalho
from src.exportador_nomus import exportar_arquivos_nomus


class AppAutomacaoBOM(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automação Engenharia CAD -> ERP Nomus")
        self.geometry("760x560")
        self.minsize(700, 500)
        self.configure(bg="#F5F7FA")

        self.var_maquina = tk.StringVar(value="TS-65")
        self.var_arq_cad = tk.StringVar()
        self.var_arq_trabalho = tk.StringVar()
        self.var_pasta_saida = tk.StringVar(value=os.path.join(BASE_DIR, "output"))

        self._construir_interface()

    def _construir_interface(self):
        # Cabeçalho
        header = tk.Frame(self, bg="#1F4E78", padx=20, pady=14)
        header.pack(fill="x")
        tk.Label(
            header,
            text="Automação BOM CAD → ERP Nomus",
            font=("Segoe UI", 15, "bold"),
            fg="white",
            bg="#1F4E78",
        ).pack(anchor="w")
        tk.Label(
            header,
            text="Geração automatizada de Importação de Produtos, Lista de Materiais e Tabela de Códigos",
            font=("Segoe UI", 9),
            fg="#D9E1F2",
            bg="#1F4E78",
        ).pack(anchor="w")

        container = tk.Frame(self, bg="#F5F7FA", padx=20, pady=15)
        container.pack(fill="both", expand=True)

        # 0. Identificação da Máquina e Pasta de Saída
        frame_cfg = tk.LabelFrame(
            container,
            text=" 1. Dados Gerais da Máquina ",
            font=("Segoe UI", 10, "bold"),
            bg="#F5F7FA",
            fg="#1F4E78",
            padx=12,
            pady=10,
        )
        frame_cfg.pack(fill="x", pady=(0, 12))

        tk.Label(frame_cfg, text="Código da Máquina (Raiz):", bg="#F5F7FA", font=("Segoe UI", 9, "bold")).grid(
            row=0, column=0, sticky="w", pady=4
        )
        tk.Entry(frame_cfg, textvariable=self.var_maquina, width=18, font=("Segoe UI", 10)).grid(
            row=0, column=1, sticky="w", padx=8, pady=4
        )
        tk.Label(
            frame_cfg,
            text="(Ex.: TS-65, TS-55 — Nome da lista será <MÁQUINA>-001)",
            bg="#F5F7FA",
            fg="#595959",
            font=("Segoe UI", 8),
        ).grid(row=0, column=2, sticky="w", pady=4)

        tk.Label(frame_cfg, text="Pasta de Saída:", bg="#F5F7FA", font=("Segoe UI", 9)).grid(
            row=1, column=0, sticky="w", pady=4
        )
        tk.Entry(frame_cfg, textvariable=self.var_pasta_saida, width=55, font=("Segoe UI", 9)).grid(
            row=1, column=1, columnspan=2, sticky="we", padx=8, pady=4
        )
        tk.Button(
            frame_cfg,
            text="Escolher Pasta...",
            command=self._selecionar_pasta_saida,
            bg="#E9EDF4",
            relief="groove",
        ).grid(row=1, column=3, padx=4, pady=4)
        frame_cfg.columnconfigure(2, weight=1)

        # Etapa 1
        frame_etapa1 = tk.LabelFrame(
            container,
            text=" 2. Etapa 1 — Destrinchar BOM do CAD e Gerar Planilha de Trabalho ",
            font=("Segoe UI", 10, "bold"),
            bg="#F5F7FA",
            fg="#1F4E78",
            padx=12,
            pady=10,
        )
        frame_etapa1.pack(fill="x", pady=(0, 12))

        tk.Label(frame_etapa1, text="Planilha BOM (CAD):", bg="#F5F7FA", font=("Segoe UI", 9)).grid(
            row=0, column=0, sticky="w", pady=4
        )
        tk.Entry(frame_etapa1, textvariable=self.var_arq_cad, width=55, font=("Segoe UI", 9)).grid(
            row=0, column=1, sticky="we", padx=8, pady=4
        )
        tk.Button(
            frame_etapa1,
            text="Selecionar BOM...",
            command=self._selecionar_bom_cad,
            bg="#E9EDF4",
            relief="groove",
        ).grid(row=0, column=2, padx=4, pady=4)

        btn_etapa1 = tk.Button(
            frame_etapa1,
            text="▶  Etapa 1 – Gerar Planilha de Trabalho",
            command=self._executar_etapa1,
            bg="#2E75B6",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=6,
            cursor="hand2",
        )
        btn_etapa1.grid(row=1, column=0, columnspan=3, sticky="we", pady=(8, 2))
        frame_etapa1.columnconfigure(1, weight=1)

        # Etapa 2
        frame_etapa2 = tk.LabelFrame(
            container,
            text=" 3. Etapa 2 — Gerar Planilhas Finais Nomus (Produtos + Lista de Materiais + Códigos) ",
            font=("Segoe UI", 10, "bold"),
            bg="#F5F7FA",
            fg="#1F4E78",
            padx=12,
            pady=10,
        )
        frame_etapa2.pack(fill="x", pady=(0, 12))

        tk.Label(frame_etapa2, text="Planilha de Trabalho Revisada:", bg="#F5F7FA", font=("Segoe UI", 9)).grid(
            row=0, column=0, sticky="w", pady=4
        )
        tk.Entry(frame_etapa2, textvariable=self.var_arq_trabalho, width=55, font=("Segoe UI", 9)).grid(
            row=0, column=1, sticky="we", padx=8, pady=4
        )
        tk.Button(
            frame_etapa2,
            text="Selecionar Planilha...",
            command=self._selecionar_planilha_trabalho,
            bg="#E9EDF4",
            relief="groove",
        ).grid(row=0, column=2, padx=4, pady=4)

        btn_etapa2 = tk.Button(
            frame_etapa2,
            text="✔  Etapa 2 – Gerar Arquivos de Importação Nomus",
            command=self._executar_etapa2,
            bg="#385723",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=6,
            cursor="hand2",
        )
        btn_etapa2.grid(row=1, column=0, columnspan=3, sticky="we", pady=(8, 2))
        frame_etapa2.columnconfigure(1, weight=1)

        # Log de status
        tk.Label(container, text="Histórico de Execução:", bg="#F5F7FA", font=("Segoe UI", 9, "bold")).pack(
            anchor="w"
        )
        self.txt_log = tk.Text(container, height=8, font=("Consolas", 9), bg="white", state="disabled")
        self.txt_log.pack(fill="both", expand=True, pady=(4, 0))

    def _log(self, msg: str):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")
        self.update_idletasks()

    def _selecionar_pasta_saida(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta de saída")
        if pasta:
            self.var_pasta_saida.set(pasta)

    def _selecionar_bom_cad(self):
        arq = filedialog.askopenfilename(
            title="Selecione a planilha BOM exportada do CAD",
            filetypes=[("Planilhas Excel", "*.xlsx")],
        )
        if arq:
            self.var_arq_cad.set(arq)

    def _selecionar_planilha_trabalho(self):
        arq = filedialog.askopenfilename(
            title="Selecione a Planilha de Trabalho revisada",
            filetypes=[("Planilhas Excel", "*.xlsx")],
        )
        if arq:
            self.var_arq_trabalho.set(arq)

    def _executar_etapa1(self):
        arq_cad = self.var_arq_cad.get().strip()
        maquina = self.var_maquina.get().strip().upper()
        pasta_saida = self.var_pasta_saida.get().strip()

        if not maquina:
            messagebox.showwarning("Atenção", "Informe o código da máquina (ex.: TS-65).")
            return
        if not arq_cad or not os.path.isfile(arq_cad):
            messagebox.showwarning("Atenção", "Selecione um arquivo BOM do CAD válido (.xlsx).")
            return

        try:
            os.makedirs(pasta_saida, exist_ok=True)
            self._log(f"[Etapa 1] Lendo BOM do CAD: {arq_cad}...")
            config = carregar_config()
            itens_cad = ler_bom_cad(arq_cad)
            self._log(f"[Etapa 1] {len(itens_cad)} itens identificados no BOM do CAD.")

            linhas_trab = processar_bom_para_planilha_trabalho(itens_cad, maquina, config)
            qtd_alertas = sum(1 for l in linhas_trab if l.alertas)

            arq_saida_trab = os.path.join(pasta_saida, f"Planilha_Trabalho_{maquina}.xlsx")
            salvar_planilha_trabalho(linhas_trab, arq_saida_trab, maquina)

            self.var_arq_trabalho.set(arq_saida_trab)
            self._log(f"[Etapa 1] Planilha de Trabalho gerada com {len(linhas_trab)} linhas ({qtd_alertas} linhas com alertas em amarelo).")
            self._log(f"          Salva em: {arq_saida_trab}")

            abrir = messagebox.askyesno(
                "Etapa 1 Concluída",
                f"Planilha de Trabalho gerada com sucesso!\n\n"
                f"• Total de linhas: {len(linhas_trab)}\n"
                f"• Linhas com alerta em amarelo para revisão: {qtd_alertas}\n\n"
                f"Arquivo: {arq_saida_trab}\n\n"
                f"Deseja abrir a Planilha de Trabalho no Excel agora?",
            )
            if abrir:
                os.startfile(arq_saida_trab)
        except Exception as e:
            self._log(f"[ERRO Etapa 1] {e}")
            messagebox.showerror("Erro na Etapa 1", f"Ocorreu um erro ao processar o BOM:\n\n{e}")

    def _executar_etapa2(self):
        arq_trab = self.var_arq_trabalho.get().strip()
        maquina = self.var_maquina.get().strip().upper()
        pasta_saida = self.var_pasta_saida.get().strip()

        if not maquina:
            messagebox.showwarning("Atenção", "Informe o código da máquina (ex.: TS-65).")
            return
        if not arq_trab or not os.path.isfile(arq_trab):
            messagebox.showwarning("Atenção", "Selecione a Planilha de Trabalho revisada (.xlsx).")
            return

        try:
            self._log(f"[Etapa 2] Lendo Planilha de Trabalho: {arq_trab}...")
            config = carregar_config()
            linhas = ler_planilha_trabalho(arq_trab, config)

            arq_prod, arq_bom, arq_tab = exportar_arquivos_nomus(
                linhas, maquina, pasta_saida, config=config
            )
            self._log("[Etapa 2] Arquivos gerados com sucesso:")
            self._log(f"   1. Produtos: {arq_prod}")
            self._log(f"   2. Lista de Materiais: {arq_bom}")
            self._log(f"   3. Tabela de Códigos: {arq_tab}")

            abrir = messagebox.askyesno(
                "Etapa 2 Concluída",
                f"Os 3 arquivos finais foram gerados com sucesso na pasta:\n{pasta_saida}\n\n"
                f"1. {os.path.basename(arq_prod)}\n"
                f"2. {os.path.basename(arq_bom)}\n"
                f"3. {os.path.basename(arq_tab)}\n\n"
                f"Deseja abrir a pasta de saída agora?",
            )
            if abrir:
                os.startfile(pasta_saida)
        except Exception as e:
            self._log(f"[ERRO Etapa 2] {e}")
            messagebox.showerror("Erro na Etapa 2", f"Ocorreu um erro ao gerar os arquivos finais:\n\n{e}")


if __name__ == "__main__":
    app = AppAutomacaoBOM()
    app.mainloop()
