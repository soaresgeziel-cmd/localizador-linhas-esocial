import os
import re
import threading
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd

# Configuração visual do tema "Windows Light"
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

EVENTOS_DICI = {
    "-- Selecione um Evento --": "",
    "S-1000 - Info. Empregador/Órgão Público": "S-1000",
    "S-1005 - Tabela de Estabelecimentos/Obras": "S-1005",
    "S-1010 - Tabela de Rubricas": "S-1010",
    "S-1020 - Tabela de Lotações Tributárias": "S-1020",
    "S-1070 - Tabela de Processos Adm./Judiciais": "S-1070",
    "S-1200 - Remuneração RGPS": "S-1200",
    "S-1202 - Remuneração RPPS": "S-1202",
    "S-1207 - Benefícios - Entes Públicos": "S-1207",
    "S-1210 - Pagamentos de Rendimentos": "S-1210",
    "S-1260 - Prod. Rural Pessoa Física": "S-1260",
    "S-1270 - Avulsos Não Portuários": "S-1270",
    "S-1280 - Info. Complementares": "S-1280",
    "S-1298 - Reabertura de Periódicos": "S-1298",
    "S-1299 - Fechamento de Periódicos": "S-1299",
    "S-2190 - Admissão Preliminar": "S-2190",
    "S-2200 - Cadastramento / Admissão": "S-2200",
    "S-2205 - Alteração Cadastral": "S-2205",
    "S-2206 - Alteração Contratual": "S-2206",
    "S-2210 - Comunicação Acidente Trabalho (CAT)": "S-2210",
    "S-2220 - Monitoramento Saúde (ASO)": "S-2220",
    "S-2230 - Afastamento Temporário": "S-2230",
    "S-2240 - Condições Ambientais (SST)": "S-2240",
    "S-2298 - Reintegração": "S-2298",
    "S-2299 - Desligamento": "S-2299",
    "S-2300 - Trabalhador Sem Vínculo - Início": "S-2300",
    "S-2306 - Sem Vínculo - Alteração": "S-2306",
    "S-2399 - Sem Vínculo - Término": "S-2399",
    "S-2400 - Cadastro Beneficiário": "S-2400",
    "S-2500 - Processo Trabalhista": "S-2500",
    "S-2501 - Tributos Processo Trabalhista": "S-2501",
    "S-3000 - Exclusão de Eventos": "S-3000",
}

EVENTOS_COM_COMPETENCIA = ["S-1200", "S-1202", "S-1207", "S-1210"]


class AppLocalizadorLinhas(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("Localizador de Linhas eSocial")
        self.geometry("740x620")
        self.resizable(False, False)
        self.configure(fg_color="#F8F9FA")

        self.caminho_pasta_origem = ""
        self.caminho_pasta_destino = os.path.join(os.path.expanduser("~"), "Downloads")
        self.caminho_excel_lote = ""
        self.total_termos_lote = 0

        # --- CABEÇALHO COMPACTO ---
        self.lbl_titulo = ctk.CTkLabel(
            self,
            text="🔎 Localizador de Linhas - eSocial",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#005A9E",
        )
        self.lbl_titulo.pack(pady=(10, 2))

        # --- CARD 1: ORIGEM DOS DADOS ---
        self.card1 = ctk.CTkFrame(
            self, corner_radius=6, fg_color="#FFFFFF", border_width=1, border_color="#D1D1D1"
        )
        self.card1.pack(fill="x", padx=15, pady=3)

        self.lbl_c1 = ctk.CTkLabel(
            self.card1,
            text="1. Origem dos Dados",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#005A9E",
        )
        self.lbl_c1.grid(row=0, column=0, columnspan=3, padx=12, pady=(4, 2), sticky="w")

        self.lbl_evento = ctk.CTkLabel(self.card1, text="Evento:", text_color="#333333", font=ctk.CTkFont(size=12))
        self.lbl_evento.grid(row=1, column=0, padx=(12, 5), pady=4, sticky="w")

        self.combo_evento = ctk.CTkComboBox(
            self.card1,
            values=list(EVENTOS_DICI.keys()),
            width=340,
            command=self.ao_mudar_evento,
            font=ctk.CTkFont(size=12),
        )
        self.combo_evento.set("-- Selecione um Evento --")
        self.combo_evento.grid(row=1, column=1, padx=5, pady=4, sticky="w")

        self.btn_pasta_origem = ctk.CTkButton(
            self.card1,
            text="📁 Selecionar Pasta Mãe",
            fg_color="#0078D4",
            hover_color="#005A9E",
            text_color="#FFFFFF",
            font=ctk.CTkFont(size=12),
            height=28,
            command=self.selecionar_pasta_origem,
        )
        self.btn_pasta_origem.grid(row=1, column=2, padx=12, pady=4)

        self.lbl_pasta_origem_status = ctk.CTkLabel(
            self.card1,
            text="⚠️ Selecione a pasta onde estão os arquivos (ou subpastas)",
            text_color="#D83B01",
            font=ctk.CTkFont(size=11),
        )
        self.lbl_pasta_origem_status.grid(
            row=2, column=0, columnspan=3, padx=12, pady=(0, 4), sticky="w"
        )

        # --- CARD 2: FILTROS DE OPERAÇÃO E COMPETÊNCIA ---
        self.card2 = ctk.CTkFrame(
            self, corner_radius=6, fg_color="#FFFFFF", border_width=1, border_color="#D1D1D1"
        )
        self.card2.pack(fill="x", padx=15, pady=3)

        self.lbl_c2 = ctk.CTkLabel(
            self.card2,
            text="2. Regras de Filtragem",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#005A9E",
        )
        self.lbl_c2.grid(row=0, column=0, columnspan=4, padx=12, pady=(4, 2), sticky="w")

        self.lbl_operacao = ctk.CTkLabel(self.card2, text="Operação:", text_color="#333333", font=ctk.CTkFont(size=12))
        self.lbl_operacao.grid(row=1, column=0, padx=(12, 5), pady=4, sticky="w")

        self.combo_operacao = ctk.CTkComboBox(
            self.card2, values=["Normal", "Retificação", "Ambas"], width=120, font=ctk.CTkFont(size=12)
        )
        self.combo_operacao.grid(row=1, column=1, padx=5, pady=4, sticky="w")

        self.lbl_comp = ctk.CTkLabel(self.card2, text="Competência(s):", text_color="#333333", font=ctk.CTkFont(size=12))
        self.lbl_comp.grid(row=1, column=2, padx=(15, 5), pady=4, sticky="w")

        self.entry_comp = ctk.CTkEntry(
            self.card2, placeholder_text="Ex: 2022-10, 2022-11 (vazio = todas)", width=230, font=ctk.CTkFont(size=12)
        )
        self.entry_comp.grid(row=1, column=3, padx=(0, 12), pady=4, sticky="w")

        # --- CARD 3: DEFINIÇÃO DE ALVO ---
        self.card3 = ctk.CTkFrame(
            self, corner_radius=6, fg_color="#FFFFFF", border_width=1, border_color="#D1D1D1"
        )
        self.card3.pack(fill="x", padx=15, pady=3)

        self.lbl_c3 = ctk.CTkLabel(
            self.card3,
            text="3. Alvo da Localização",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#005A9E",
        )
        self.lbl_c3.grid(row=0, column=0, columnspan=3, padx=12, pady=(4, 2), sticky="w")

        self.var_tipo_busca = ctk.StringVar(value="CPF")

        self.radio_cpf = ctk.CTkRadioButton(
            self.card3,
            text="CPF (11 dígitos)",
            variable=self.var_tipo_busca,
            value="CPF",
            command=self.atualizar_placeholder_alvo,
            text_color="#333333",
            font=ctk.CTkFont(size=12),
        )
        self.radio_cpf.grid(row=1, column=0, padx=(12, 10), pady=4)

        self.radio_mat = ctk.CTkRadioButton(
            self.card3,
            text="Matrícula (eSocial)",
            variable=self.var_tipo_busca,
            value="Matrícula",
            command=self.atualizar_placeholder_alvo,
            text_color="#333333",
            font=ctk.CTkFont(size=12),
        )
        self.radio_mat.grid(row=1, column=1, padx=10, pady=4)

        self.entry_termo = ctk.CTkEntry(
            self.card3,
            placeholder_text="Digite o CPF (ex: 69094586649)...",
            width=310,
            font=ctk.CTkFont(size=12),
        )
        self.entry_termo.grid(
            row=2, column=0, columnspan=2, padx=(12, 10), pady=(0, 4), sticky="w"
        )

        self.btn_excel = ctk.CTkButton(
            self.card3,
            text="📊 Carregar Lote Excel",
            fg_color="#F3F3F3",
            hover_color="#E5E5E5",
            border_width=1,
            border_color="#0078D4",
            text_color="#0078D4",
            font=ctk.CTkFont(size=12),
            height=28,
            command=self.carregar_excel,
        )
        self.btn_excel.grid(row=2, column=2, padx=12, pady=(0, 4), sticky="w")

        self.lbl_badge_excel = ctk.CTkLabel(
            self.card3,
            text="",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#107C41",
        )
        self.lbl_badge_excel.grid(
            row=3, column=0, columnspan=3, padx=12, pady=(0, 2), sticky="w"
        )

        # --- CARD 4: PASTA DE DESTINO ---
        self.card4 = ctk.CTkFrame(
            self, corner_radius=6, fg_color="#FFFFFF", border_width=1, border_color="#D1D1D1"
        )
        self.card4.pack(fill="x", padx=15, pady=3)

        self.lbl_c4 = ctk.CTkLabel(
            self.card4,
            text="4. Pasta de Destino dos Arquivos",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#005A9E",
        )
        self.lbl_c4.grid(row=0, column=0, columnspan=2, padx=12, pady=(4, 2), sticky="w")

        self.btn_pasta_destino = ctk.CTkButton(
            self.card4,
            text="📂 Alterar Destino",
            fg_color="#F3F3F3",
            hover_color="#E5E5E5",
            border_width=1,
            border_color="#0078D4",
            text_color="#0078D4",
            font=ctk.CTkFont(size=12),
            height=26,
            command=self.selecionar_pasta_destino,
        )
        self.btn_pasta_destino.grid(row=1, column=0, padx=12, pady=(0, 6), sticky="w")

        self.lbl_pasta_destino_status = ctk.CTkLabel(
            self.card4,
            text=f"✔ Salvar em: {self.caminho_pasta_destino}",
            text_color="#107C41",
            font=ctk.CTkFont(size=11),
        )
        self.lbl_pasta_destino_status.grid(
            row=1, column=1, padx=10, pady=(0, 6), sticky="w"
        )

        # --- AÇÃO E STATUS ---
        self.btn_iniciar = ctk.CTkButton(
            self,
            text="🔍 INICIAR LOCALIZAÇÃO",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=38,
            fg_color="#0078D4",
            hover_color="#005A9E",
            text_color="#FFFFFF",
            state="disabled",
            command=self.iniciar_thread,
        )
        self.btn_iniciar.pack(fill="x", padx=15, pady=(8, 4))

        self.progresso = ctk.CTkProgressBar(self, width=710, progress_color="#0078D4")
        self.progresso.pack(padx=15, pady=2)
        self.progresso.set(0)

        self.lbl_status = ctk.CTkLabel(
            self,
            text="Aguardando a seleção do evento e da pasta para iniciar.",
            font=ctk.CTkFont(size=11),
            text_color="#505050",
        )
        self.lbl_status.pack(pady=(2, 6))

        self.ao_mudar_evento(self.combo_evento.get())

    def verificar_consistencia_caminho(self, codigo_evento_selecionado):
        """Mapeia se o caminho da pasta contém algum código de evento e se bate com o selecionado."""
        if not self.caminho_pasta_origem:
            return True, "⚠️ Selecione a pasta onde estão os arquivos (ou subpastas)"

        caminho_lower = self.caminho_pasta_origem.lower()
        evento_detectado = None

        # Procura por qualquer evento cadastrado no caminho da pasta
        for cod in EVENTOS_DICI.values():
            if not cod:
                continue
            cod_limpo = cod.lower()  # Ex: "s-1200"
            cod_sem_traco = cod_limpo.replace("-", "")  # Ex: "s1200"

            # Usa regex com fronteira de caracteres/palavras para evitar falsos positivos
            padrao = rf"(?:^|[\/\_\\\s\.\-])({cod_limpo}|{cod_sem_traco})(?:$|[\/\_\\\s\.\-])"
            if re.search(padrao, caminho_lower):
                evento_detectado = cod
                break

        # Se encontrou um código de evento no caminho, valida com a seleção
        if evento_detectado:
            if evento_detectado == codigo_evento_selecionado:
                return True, f"✔ Pasta compatível com o evento {codigo_evento_selecionado}: {self.caminho_pasta_origem}"
            else:
                return False, f"❌ DIVERGÊNCIA: O caminho é do {evento_detectado}, mas você selecionou {codigo_evento_selecionado}!"
        
        # Se não há nome de evento no caminho da pasta, libera leitura normal
        return True, f"✔ Pasta Mãe selecionada: {self.caminho_pasta_origem}"

    def validar_pronto_para_iniciar(self):
        evento_selecionado = EVENTOS_DICI.get(self.combo_evento.get(), "")
        
        valido_caminho, mensagem_status = self.verificar_consistencia_caminho(evento_selecionado)

        if self.caminho_pasta_origem:
            cor_texto = "#107C41" if valido_caminho else "#D83B01"
            self.lbl_pasta_origem_status.configure(text=mensagem_status, text_color=cor_texto)

        if evento_selecionado and self.caminho_pasta_origem and self.caminho_pasta_destino and valido_caminho:
            self.btn_iniciar.configure(state="normal")
            self.lbl_status.configure(text="Pronto para iniciar a localização.")
        else:
            self.btn_iniciar.configure(state="disabled")

    def atualizar_placeholder_alvo(self):
        tipo = self.var_tipo_busca.get()
        if tipo == "CPF":
            self.entry_termo.configure(
                placeholder_text="Digite o CPF (ex: 69094586649)..."
            )
        else:
            self.entry_termo.configure(
                placeholder_text="Digite a Matrícula (ex: 10000149_3_1261_6)..."
            )

    def ao_mudar_evento(self, escolha):
        codigo_evento = EVENTOS_DICI.get(escolha, "")
        if codigo_evento in EVENTOS_COM_COMPETENCIA:
            self.entry_comp.configure(
                state="normal",
                fg_color="#FFFFFF",
                text_color="#000000",
                placeholder_text="Ex: 2022-10, 2022-11 (vazio = todas)",
            )
        else:
            self.entry_comp.delete(0, "end")
            self.entry_comp.configure(
                state="disabled",
                fg_color="#C0C0C0",
                text_color="#707070",
                placeholder_text="[Não aplicável]",
            )
        self.validar_pronto_para_iniciar()

    def selecionar_pasta_origem(self):
        caminho = filedialog.askdirectory(title="Selecione a Pasta Mãe dos Arquivos de Origem")
        if caminho:
            self.caminho_pasta_origem = caminho
            self.validar_pronto_para_iniciar()

    def selecionar_pasta_destino(self):
        caminho = filedialog.askdirectory(title="Selecione a Pasta onde os Resultados Serão Salvos")
        if caminho:
            self.caminho_pasta_destino = caminho
            self.lbl_pasta_destino_status.configure(
                text=f"✔ Salvar em: {caminho}", text_color="#107C41"
            )
            self.validar_pronto_para_iniciar()

    def carregar_excel(self):
        caminho = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if caminho:
            try:
                df_lote = pd.read_excel(caminho, header=None, dtype=str)
                termos = df_lote[0].dropna().astype(str).str.strip().tolist()
                self.total_termos_lote = len(termos)
                self.caminho_excel_lote = caminho

                nome_arquivo = os.path.basename(caminho)
                self.lbl_badge_excel.configure(
                    text=f"✔ Lote Excel Carregado: {nome_arquivo} ({self.total_termos_lote} itens)"
                )
                self.entry_termo.delete(0, "end")
                self.entry_termo.configure(
                    placeholder_text="Lote em Excel ativo (campo manual ignorado)"
                )
            except Exception as e:
                messagebox.showerror(
                    "Erro ao carregar Excel", f"Não foi possível ler o arquivo: {e}"
                )

    def iniciar_thread(self):
        evento_codigo = EVENTOS_DICI.get(self.combo_evento.get(), "")
        if not evento_codigo:
            messagebox.showwarning("Atenção", "Selecione um evento válido antes de iniciar!")
            return

        if not self.caminho_pasta_origem:
            messagebox.showwarning(
                "Atenção", "Selecione a pasta mãe onde estão os arquivos de origem!"
            )
            return

        if not self.caminho_pasta_destino:
            messagebox.showwarning(
                "Atenção", "Selecione a pasta de destino dos arquivos gerados!"
            )
            return

        self.btn_iniciar.configure(state="disabled")
        self.progresso.configure(mode="indeterminated")
        self.progresso.start()
        self.lbl_status.configure(text="Varrendo pastas e arquivos de alta velocidade...")

        thread = threading.Thread(target=self.executar_busca)
        thread.daemon = True
        thread.start()

    def formatar_termo(self, termo, tipo_alvo):
        """Formata o termo ajustando zeros à esquerda para CPF se necessário."""
        termo_limpo = str(termo).strip()
        if tipo_alvo == "CPF":
            apenas_num = re.sub(r'\D', '', termo_limpo)
            if apenas_num and len(apenas_num) <= 11:
                return apenas_num.zfill(11)
        return termo_limpo

    def executar_busca(self):
        try:
            evento_desc = self.combo_evento.get()
            codigo_evento = EVENTOS_DICI[evento_desc]
            operacao_escolhida = self.combo_operacao.get()
            tipo_alvo = self.var_tipo_busca.get()
            comp_texto = self.entry_comp.get().strip()

            termos_raw = []
            prefixo_nome_arquivo = ""

            if self.caminho_excel_lote:
                df_lote = pd.read_excel(self.caminho_excel_lote, header=None, dtype=str)
                termos_raw = df_lote[0].dropna().astype(str).str.strip().tolist()
                prefixo_nome_arquivo = f"Lote_{len(termos_raw)}_itens"
            else:
                termo_dig = self.entry_termo.get().strip()
                if termo_dig:
                    termos_raw = [termo_dig]
                    prefixo_nome_arquivo = re.sub(r'[^\w\-]', '_', termo_dig)

            if not termos_raw:
                self.finalizar(
                    f"Erro: Digite um {tipo_alvo} ou carregue uma planilha Excel.",
                    sucesso=False,
                )
                return

            termos_validos = [self.formatar_termo(t, tipo_alvo) for t in termos_raw if str(t).strip()]
            termos_validos = list(set(filter(None, termos_validos)))

            if not termos_validos:
                self.finalizar("Nenhum termo válido foi informado.", sucesso=False)
                return

            # --- PRE-FILTRO SUPER-RÁPIDO EM C (Contém direto no texto) ---
            padroes_termos = [re.escape(t) for t in termos_validos]
            regex_termos_comb = re.compile(rf"(?:^|[\"',;\s])(?:{'|'.join(padroes_termos)})(?:[\"',;\s]|$)")

            # Competências
            regex_comp_comb = None
            if comp_texto and codigo_evento in EVENTOS_COM_COMPETENCIA:
                comps_list = [c.strip() for c in comp_texto.split(",") if c.strip()]
                padroes_comp = []
                for c in comps_list:
                    padroes_comp.append(re.escape(c.replace("/", "-")))
                    padroes_comp.append(re.escape(c.replace("-", "/")))
                if padroes_comp:
                    regex_comp_comb = re.compile(rf"(?:^|,|\"|;)(?:{'|'.join(padroes_comp)})(?:,|\"|;|$)")

            # Mapeamento recursivo
            arquivos = []
            for raiz, _, arquivos_pasta in os.walk(self.caminho_pasta_origem):
                for f in arquivos_pasta:
                    if f.lower().endswith((".txt", ".csv")):
                        arquivos.append(os.path.join(raiz, f))

            if not arquivos:
                self.finalizar(
                    "Nenhum arquivo .txt ou .csv foi localizado na pasta mãe.",
                    sucesso=False,
                )
                return

            linhas_normal = []
            linhas_retificacao = []

            # --- PROCESSAMENTO ULTRARRÁPIDO OTIMIZADO ---
            for arq in arquivos:
                with open(arq, "r", encoding="utf-8", errors="ignore") as f:
                    for linha in f:
                        linha_orig = linha.strip()
                        if not linha_orig:
                            continue

                        # ATALHO VELOCIDADE #1: Se o CPF/Matrícula nem aparece na string, ignora de cara!
                        if not any(t in linha_orig for t in termos_validos):
                            continue

                        # ATALHO VELOCIDADE #2: split rápido de no máximo 2 partes
                        partes = linha_orig.split(",", 2)
                        if len(partes) < 2:
                            continue

                        tipo_op = partes[1].replace('"', "").replace("'", "").strip()
                        is_norm = (tipo_op == "1")
                        is_ret = (tipo_op == "2")

                        if operacao_escolhida == "Normal" and not is_norm:
                            continue
                        if operacao_escolhida == "Retificação" and not is_ret:
                            continue
                        if not is_norm and not is_ret:
                            continue

                        # Competência
                        if regex_comp_comb and not regex_comp_comb.search(linha_orig[:200]):
                            continue

                        # Validação exata por fronteiras de caracteres
                        if regex_termos_comb.search(linha_orig):
                            if is_norm:
                                linhas_normal.append(linha_orig)
                            elif is_ret:
                                linhas_retificacao.append(linha_orig)

            # Gravação do resultado
            resumos = []
            total = 0

            if operacao_escolhida in ["Normal", "Ambas"]:
                qtd = len(linhas_normal)
                total += qtd
                resumos.append(f"Normal: {qtd}")
                if qtd > 0:
                    nome_arq = f"{prefixo_nome_arquivo}_{codigo_evento}_Normal.csv"
                    pd.DataFrame(linhas_normal, columns=[""]).to_csv(
                        os.path.join(self.caminho_pasta_destino, nome_arq),
                        index=False,
                        header=False,
                    )

            if operacao_escolhida in ["Retificação", "Ambas"]:
                qtd = len(linhas_retificacao)
                total += qtd
                resumos.append(f"Retificação: {qtd}")
                if qtd > 0:
                    nome_arq = f"{prefixo_nome_arquivo}_{codigo_evento}_Retificacao.csv"
                    pd.DataFrame(linhas_retificacao, columns=[""]).to_csv(
                        os.path.join(self.caminho_pasta_destino, nome_arq),
                        index=False,
                        header=False,
                    )

            detalhes = " | ".join(resumos)
            msg = f"Varredura Finalizada! Total localizado: {total} ({detalhes}). Salvo na pasta indicada."
            self.finalizar(msg, sucesso=True)

        except Exception as e:
            self.finalizar(f"Erro ao processar busca: {e}", sucesso=False)

    def finalizar(self, mensagem, sucesso=True):
        self.progresso.stop()
        self.progresso.configure(mode="determinate")
        self.progresso.set(1 if sucesso else 0)
        self.btn_iniciar.configure(state="normal")
        self.lbl_status.configure(text=mensagem)

        if sucesso:
            messagebox.showinfo("Resultado", mensagem)
        else:
            messagebox.showerror("Erro", mensagem)


if __name__ == "__main__":
    app = AppLocalizadorLinhas()
    app.mainloop()