import customtkinter as ctk
import tkinter.filedialog as fd
import os
from PIL import Image
import threading
from acesso_microsiga import MicrosigaAutomacao
import sys
from datetime import datetime
import pandas as pd
import re


def resource_path(relative_path):
    """Retorna o caminho absoluto para o recurso, funcionando tanto no modo de desenvolvimento quanto no PyInstaller"""
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# ===> Classe: Recebimento dos dados do usuário para seguir com a automação <=== #
class DadosUsuario:
    def __init__(self):
        self.usuario_login = ""
        self.pwd_login = ""
        self.caminho_arquivo = ""
        self.grupo_email = ""


# ===> Classe: Principal do aplicativo, interface com o usuário <=== #
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        # ---> Armazena o tempo de início da aplicação <---
        self.start_time = datetime.now()
        self.title("Cadastro de Previsão de Vendas - Microsiga")
        self.geometry("1300x550")
        self.minsize(800, 550)
        ctk.set_appearance_mode("system")

        self.dados = DadosUsuario()
        self.evento_parar = threading.Event()
        self.automation_thread = None

        # Configura o grid principal para ser responsivo
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.criar_widgets()

    # ===> Cria e posiciona todos os componentes visuais da janela <=== #
    def criar_widgets(self):

        # --- Frame Principal ---
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)

        # --- Fontes ---
        labels_font = ctk.CTkFont(family="Arial", size=12, weight="bold")
        entry_font = ctk.CTkFont(family="Arial", size=12)
        btn_font = ctk.CTkFont(family="Arial", size=14, weight="bold")

        # --- Frame de Entradas (Email, Usuário, Senha) ---
        inputs_frame = ctk.CTkFrame(main_frame)
        inputs_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        inputs_frame.grid_columnconfigure(1, weight=1)

        # --- E-mail ---
        ctk.CTkLabel(
            inputs_frame, text="E-mail para Notificação:", font=labels_font
        ).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.entry_app_email = ctk.CTkEntry(
            inputs_frame,
            font=entry_font,
            placeholder_text="email@suaempresa.com; outro@suaempresa.com",
        )
        self.entry_app_email.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.entry_app_email.insert(
            0,
            "",  # ATENÇÃO: Deixado em branco para o usuário preencher.
        )

        # --- Usuário ---
        ctk.CTkLabel(inputs_frame, text="Usuário Microsiga:", font=labels_font).grid(
            row=1, column=0, padx=10, pady=10, sticky="w"
        )
        self.entry_app_user = ctk.CTkEntry(
            inputs_frame, placeholder_text="Informe seu usuário", font=entry_font
        )
        self.entry_app_user.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        # --- Senha ---
        ctk.CTkLabel(inputs_frame, text="Senha Microsiga:", font=labels_font).grid(
            row=2, column=0, padx=10, pady=10, sticky="w"
        )
        self.entry_app_pwd = ctk.CTkEntry(
            inputs_frame,
            placeholder_text="Informe sua senha",
            show="*",
            font=entry_font,
        )
        self.entry_app_pwd.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        self.check_mostrar_pwd = ctk.CTkCheckBox(
            inputs_frame,
            text="Mostrar senha",
            font=entry_font,
            command=self.mostrar_pwd,
        )
        self.check_mostrar_pwd.grid(row=2, column=2, padx=10, pady=10, sticky="w")

        # --- Frame do Modelo e Arquivo ---
        file_frame = ctk.CTkFrame(main_frame)
        file_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=10)
        file_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            file_frame, text="Abaixo modelo de arquivo .xlsx:", font=labels_font
        ).grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")

        # --- Imagem ---
        try:
            # MODIFICADO: Nome do arquivo de imagem genérico
            caminho_imagem = resource_path(os.path.join("img", "modelo_planilha.png"))

            imagem_modelo = ctk.CTkImage(
                light_image=Image.open(caminho_imagem), size=(1262, 164)
            )
            label_imagem = ctk.CTkLabel(file_frame, image=imagem_modelo, text="")
            label_imagem.grid(
                row=1, column=0, columnspan=2, padx=10, pady=5, sticky="ew"
            )
        except Exception as e:
            msg_erro_img = (
                f"Aviso: Imagem 'modelo_planilha.png' não encontrada. (Erro: {e})"
            )
            ctk.CTkLabel(file_frame, text=msg_erro_img, text_color="orange").grid(
                row=1, column=0, columnspan=2, padx=10, pady=5, sticky="w"
            )

        # --- Seleção de Arquivo ---
        ctk.CTkLabel(file_frame, text="Arquivo:", font=labels_font).grid(
            row=2, column=0, padx=10, pady=(15, 10), sticky="w"
        )
        self.label_arquivo_xlsx = ctk.CTkLabel(
            file_frame,
            text="Nenhum arquivo selecionado",
            font=entry_font,
            text_color="gray",
        )
        self.label_arquivo_xlsx.grid(
            row=2, column=0, padx=(80, 10), pady=(15, 10), sticky="w"
        )
        self.btn_localizar_arquivo = ctk.CTkButton(
            file_frame,
            text="Buscar Arquivo",
            font=btn_font,
            command=self.receber_arquivo,
        )
        self.btn_localizar_arquivo.grid(
            row=2, column=1, padx=10, pady=(15, 10), sticky="e"
        )

        # --- Frame de Ações e Status ---
        action_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        action_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(20, 0))
        action_frame.grid_columnconfigure(0, weight=1)

        self.btn_iniciar = ctk.CTkButton(
            action_frame,
            text="Cadastrar Previsões",
            font=btn_font,
            height=40,
            fg_color="#4CAF50",
            hover_color="#45a049",
            command=self.iniciar_automacao,
        )
        self.btn_iniciar.grid(row=0, column=1, padx=5, sticky="e")

        self.btn_parar = ctk.CTkButton(
            action_frame,
            text="Parar Execução",
            font=btn_font,
            height=40,
            fg_color="#f44336",
            hover_color="#da190b",
            command=self.parar_automacao,
            state="disabled",
        )
        self.btn_parar.grid(row=0, column=2, padx=5, sticky="e")

        self.return_user = ctk.CTkLabel(
            action_frame, text="Pronto para iniciar.", font=entry_font
        )
        self.return_user.grid(row=0, column=0, padx=10, sticky="w")

    # ===> Função: Mostrar senha ou não para o usuário <=== #
    def mostrar_pwd(self):
        self.entry_app_pwd.configure(
            show="" if self.entry_app_pwd.cget("show") == "*" else "*"
        )

    # ===> Função para recebimento do arquivo dentro da interface <=== #
    def receber_arquivo(self):
        caminho = fd.askopenfilename(
            title="Selecionar arquivo .xlsx",
            filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")),
        )
        if caminho:
            self.dados.caminho_arquivo = caminho
            self.label_arquivo_xlsx.configure(
                text=os.path.basename(caminho), text_color="green"
            )
            print(f"Arquivo selecionado: {self.dados.caminho_arquivo}")

    # ===> Função: Cancelamento da automação a pedido do usuário <=== #
    def parar_automacao(self):
        if self.automation_thread and self.automation_thread.is_alive():
            self.evento_parar.set()
            self.atualizar_status("Sinal de parada enviado...", "orange")

    # ===> Função: Identificar último registro de log
    @staticmethod
    def encontra_ultimo_arquivo(folder_path, base_name):
        arquivos = os.listdir(folder_path)
        arquivos_data = [
            file for file in arquivos if re.match(rf"{base_name}(\d+)\.csv", file)
        ]
        if not arquivos_data:
            return None
        ultimo_arquivo = max(
            arquivos_data,
            key=lambda x: (
                int(re.search(r"\d+", x).group()) if re.search(r"\d+", x) else 0
            ),
        )
        return ultimo_arquivo

    # ===> Função: Cria novo registro de log
    @staticmethod
    def cria_proximo_arquivo(folder_path, base_name):
        ultimo_arquivo = App.encontra_ultimo_arquivo(folder_path, base_name)
        if ultimo_arquivo:
            ultimo_numero = int(re.search(r"\d+", ultimo_arquivo).group())
            proximo_numero = ultimo_numero + 1
        else:
            proximo_numero = 1
        nome_proximo_arquivo = f"{base_name}{proximo_numero}.csv"
        proximo_caminho = os.path.join(folder_path, nome_proximo_arquivo)
        return proximo_caminho

    # ===> Função: Iniciar nossa automação, com ela podemos inciar a captação dos dados do usuário e também informar se falta algo essencial que impede o funcionamento correto da automação
    def iniciar_automacao(self):
        self.dados.usuario_login = self.entry_app_user.get()
        self.dados.pwd_login = self.entry_app_pwd.get()
        self.dados.grupo_email = self.entry_app_email.get()

        erros = []
        if not self.dados.usuario_login:
            erros.append("Usuário é obrigatório.")
        if not self.dados.pwd_login:
            erros.append("Senha é obrigatória.")
        if not self.dados.caminho_arquivo:
            erros.append("É necessário selecionar um arquivo.")
        if not self.dados.grupo_email:
            erros.append("E-mail é obrigatório.")
        else:
            # ATENÇÃO: A validação de domínio foi alterada para um placeholder.
            # Você pode ajustar o domínio ou remover esta validação se não for necessária.
            emails = self.dados.grupo_email.split(";")
            for email in emails:
                email_limpo = email.strip()
                if email_limpo and not email_limpo.endswith("@suaempresa.com"):
                    erros.append(
                        f"O e-mail '{email_limpo}' não é um domínio válido (@suaempresa.com)."
                    )
                    break

        if erros:
            mensagem = "Erros de Validação:\n- " + "\n- ".join(erros)
            self.atualizar_status(mensagem, "yellow")
            return

        # Lê o arquivo para obter a contagem de itens
        quantidade_itens = 0
        try:
            df = pd.read_excel(self.dados.caminho_arquivo)
            quantidade_itens = len(df.index)
        except Exception as e:
            print(f"Aviso: Não foi possível ler o arquivo Excel para o log. Erro: {e}")

        # Configura a UI para o estado de execução
        self.btn_iniciar.configure(state="disabled")
        self.btn_parar.configure(state="normal")
        self.evento_parar.clear()
        self.atualizar_status("Iniciando automação em segundo plano...", "cyan")

        # Inicia a thread de automação
        automator = MicrosigaAutomacao(self.dados, self.evento_parar)
        self.automation_thread = threading.Thread(
            target=lambda: self.executar_e_atualizar_ui(automator), daemon=True
        )
        self.automation_thread.start()

    # ===> Função: Executa a automação e agenda a atualização da UI para a thread principal
    def executar_e_atualizar_ui(self, automator):
        resultado = automator.executar()
        self.after(0, self.finalizar_automacao, resultado)

    # ===> Função: Sempre executada na thread principal, sendo segura para a UI
    def finalizar_automacao(self, resultado):
        if self.evento_parar.is_set():
            self.atualizar_status("Automação interrompida pelo usuário.", "orange")
        elif resultado and resultado.get("sucesso"):
            self.atualizar_status(
                resultado.get("mensagem", "Automação concluída com sucesso!"), "green"
            )
        else:
            msg_falha = resultado.get(
                "mensagem", "Falha na automação. Verifique o console."
            )
            self.atualizar_status(msg_falha, "red")

        self.btn_iniciar.configure(state="normal")
        self.btn_parar.configure(state="disabled")

    # ===> Função: Centralizada para atualizar a label de status
    def atualizar_status(self, texto, cor):
        self.return_user.configure(text=texto, text_color=cor)


if __name__ == "__main__":
    app = App()
    app.mainloop()
