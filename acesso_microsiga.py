import time
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import os
import pandas as pd
from datetime import datetime, timedelta
from selenium.webdriver.common.keys import Keys
from win32com.client import Dispatch
from selenium.webdriver.common.action_chains import ActionChains

# ===> Classe: Principal onde todo o processo de automação esta alocado
class MicrosigaAutomacao:
    # ===> Função: Recebimento dos dados do usuário recebidos da interface com usuário
    def __init__(self, dados_usuario, evento_parar):
        self.login = dados_usuario.usuario_login
        self.senha = dados_usuario.pwd_login
        self.arquivo = dados_usuario.caminho_arquivo
        self.user_mail = dados_usuario.grupo_email
        self.evento_parar = evento_parar
        self.driver = None
        self.actions = None
        self.wait = None
        self.short_wait = None

    # ===> Função: Verifica se o usuário clicou no botão de cancelar o andamento da automação
    def _verificar_parada(self):
        if self.evento_parar.is_set():
            raise InterruptedError("Automação interrompida pelo usuário.")

    # ===> Função: Configurador do driver, para funcionamento da automação Web
    def _iniciar_driver(self):
        print("Iniciando o navegador.")
        chrome_options = Options()
        arguments = [
            "--lang=pt-BR",
            "--start-maximized",
            "--disable-notifications",
            "--disable-gpu",
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-extensions",
            "--log-level=3",
            # "--headless",
        ]

        for argument in arguments:
            chrome_options.add_argument(argument)
        chrome_options.add_experimental_option("detach", True)
        service = ChromeService()
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.short_wait = WebDriverWait(self.driver, 3)
        self.wait = WebDriverWait(self.driver, 45)
        self.actions = ActionChains(self.driver)

    # ===> Função: Realiza a busca do primeiro elemento que corresponde ao caminho usando querySelector, retorna um único WebElement ou None
    def _find_first_element(self, path_selectors):
        script = """
            let element = document;
            const selectors = arguments[0];
            for (let i = 0; i < selectors.length; i++) {
                if (element.shadowRoot) { element = element.shadowRoot; }
                element = element.querySelector(selectors[i]);
                if (!element) { return null; }
            }
            return element;
        """
        return self.wait.until(
            lambda driver: driver.execute_script(script, path_selectors),
            f"Timeout ao tentar encontrar o elemento no caminho Shadow DOM: {path_selectors}",
        )

    # ===> Função: Realiza a busca de todos os elementos que correspondem ao caminho usando querySelectorAll
    def _find_all_elements(self, path_selectors):
        script = """
            let elements = [document]; 
            const selectors = arguments[0];
            const last_selector_index = selectors.length - 1;
            for (let i = 0; i < last_selector_index; i++) {
                let current_elements = [];
                for (const el of elements) {
                    let root = el.shadowRoot ? el.shadowRoot : el;
                    let found_element = root.querySelector(selectors[i]);
                    if (found_element) current_elements.push(found_element);
                }
                elements = current_elements;
                if (elements.length === 0) return [];
            }
            let final_elements = [];
            const last_selector = selectors[last_selector_index];
            for (const el of elements) {
                let root = el.shadowRoot ? el.shadowRoot : el;
                final_elements.push(...root.querySelectorAll(last_selector));
            }
            return final_elements;
        """

        return self.wait.until(
            lambda driver: driver.execute_script(script, path_selectors),
            f"Timeout ao tentar encontrar a lista de elementos no caminho Shadow DOM: {path_selectors}",
        )

    # ===> Função: Tratar pop-ups - Pesquisa da rotina, reforma tributária e limite de usuários
    def tratar_pop_ups(self):
        print("\n--- INICIANDO TRATAMENTO DE POP-UPS DA ROTINA ---")
        tempo_maximo_espera = 30  # Tempo total para resolver todos os pop-ups
        inicio_espera = datetime.now()

        while (datetime.now() - inicio_espera).total_seconds() < tempo_maximo_espera:
            self._verificar_parada()

            # ===> Try: Tenta encontrar e resolver o pop-up de "reforma tributária"
            try:
                seletor_botao_fechar_popup = (By.CSS_SELECTOR, "wa-button#COMP6013")
                botao_fechar = self.short_wait.until(
                    EC.element_to_be_clickable((seletor_botao_fechar_popup))
                )
                print("---> Pop-up de reforma tributária encontrado. Fechando...")
                botao_fechar.click()
                time.sleep(1)
                print("    - Aguardando pop-up de parâmetros desaparecer...")
                self.short_wait.until(EC.staleness_of(botao_fechar))
                print("    - Pop-up fechado com sucesso.")
                continue  # Reinicia o loop para verificar o estado novamente
            except TimeoutException:
                pass  # Pop-up não encontrado, continue

            # 3. VERIFICA O POP-UP DE "LIMITE DE ACESSO":
            try:
                seletor_botao_limite_acesso = (By.CSS_SELECTOR, "wa-button#COMP4511")
                botao_limite = self.short_wait.until(
                    EC.element_to_be_clickable(seletor_botao_limite_acesso)
                )
                print(
                    "---> Pop-up de 'Limite de Acesso' visível encontrado. Iniciando recuperação..."
                )
                botao_limite.click()
                time.sleep(1)
                print("    - Aguardando pop-up de limite desaparecer...")
                self.short_wait.until(EC.staleness_of(botao_limite))

                # Clica em Home e navega para a rotina novamente
                seletor_botao_home = (By.CSS_SELECTOR, "wa-button#COMP3027")
                # Para a recuperação, podemos usar a espera mais longa para garantir
                botao_home = self.short_wait.until(
                    EC.element_to_be_clickable(seletor_botao_home)
                )
                botao_home.click()

                self.navegar_rotina(modulo="MATASC4")
                print("---> Recuperação concluída. Reiniciando verificação de pop-ups.")
                continue  # Reinicia o loop para reavaliar a nova tela.
            except TimeoutException:
                pass  # Pop-up não está visível, continue.

            print("  - Nenhum pop-up ativo detectado. Aguardando a tela carregar...")
            time.sleep(1)  # Pequena pausa para evitar sobrecarregar o processador

            # ===> Try: Tenta encontrar e fechar o pop-up de "escolha da rotina"
            try:
                seletor_incluir = (By.CSS_SELECTOR, "wa-button#COMP4566")
                self.short_wait.until(EC.element_to_be_clickable(seletor_incluir))
                print("--> Tela da rotina principal está pronta.")
                return
            except TimeoutException:
                print(
                    f"---> Não foi identificada a rotina. Verificar se há pop-ups a serem fechados."
                )
                pass  # Pop-up não encontrado, continue para a próxima verificação

        # Se o loop terminar sem encontrar a condição de sucesso, lança um erro.
        raise TimeoutException(
            "Não foi possível resolver os pop-ups e chegar na tela da rotina no tempo esperado."
        )

    # ===> Função: Identifica campo de pesquisa e realiza a pesquisa pela rotina - Previsão de vendas
    def navegar_rotina(self, modulo: str):
        try:
            time.sleep(5)
            self._verificar_parada()
            print("\n--- NAVEGANDO PARA A ROTINA 'MATASC4' ---")
            modulo = f"{modulo}"
            path_campo_pesquisa = ["wa-text-input#COMP3056", "input"]
            campo_pesquisa_real = self._find_first_element(path_campo_pesquisa)
            campo_pesquisa_real.click()
            time.sleep(0.5)
            campo_pesquisa_real.clear()
            time.sleep(0.5)
            campo_pesquisa_real.send_keys(modulo)
            time.sleep(0.5)
            campo_pesquisa_real.send_keys(Keys.CONTROL, Keys.F3)

            try:
                caminho_dos_resultados = "wa-menu-item[data-advpl='tmenuitem']"
                lista_de_resultados = self.short_wait.until(
                    EC.visibility_of_all_elements_located(
                        (By.CSS_SELECTOR, caminho_dos_resultados)
                    )
                )
                if lista_de_resultados:
                    print(
                        f"-> Foram encontrados {len(lista_de_resultados)} resultados. Selecionando o primeiro."
                    )
                    lista_de_resultados[0].click()

            except TimeoutException:
                # Se o menu não aparecer em 3 segundos, o código vem para cá e continua
                print(
                    "--> Menu de resultados não apareceu. A rotina pode ter sido acessada diretamente."
                )
            except Exception as e:
                # Captura outros erros inesperados e continua a execução
                print(
                    f"Aviso: Ocorreu um erro não crítico ao verificar o menu de resultados. Erro: {e}"
                )

        except Exception as e:
            print(
                f"---> Ocorreu um erro em clicar e preencher o campo de pesquisa.\nErro:{e}"
            )
            raise  # Lança o erro para que o bloco principal possa tratá-lo

    # ===> Função: Leitura e tratamento da planilha
    def ler_arquivo(self):
        self._verificar_parada()
        try:
            print(f"Lendo o arquivo xlsx para obter os dados: {self.arquivo}")
            df = pd.read_excel(self.arquivo, header=1)
            df_new = df.drop(
                columns=[
                    "C5_EMISSAO",
                    "C6_ITEM",
                    "C6_QTDENT",
                    "C6_PRCVEN",
                    "C6_CLASFIS",
                    "C6_CF",
                    "C6_CSTPIS",
                    "B1_ORIGEM",
                    "B1_CLASFIS",
                ],
                errors="ignore",  # Ignora se alguma coluna não existir
            )
            print(df_new)
            return df_new
        except FileNotFoundError:
            print(f"Erro: Arquivo não encontrado em {self.arquivo}")
            return None
        except Exception as e:
            print(f"Erro ao ler o arquivo: {e}")
            return None

    # ===> Função:Tratar os dados remanescentes da planilha e realizar o processo de inclusão dos dados da planilha na rotina
    def processar_provisoes(self, df):

        if df is None:
            return [], [], []
        print("Iniciando processamento das provisões na rotina: MATASC4.")
        pedidos_sucesso = []
        pedidos_falha = set()
        pn_voss_falha = set()

        # ---> Define os seletores
        seletor_cod_cliente = ["wa-text-input#COMP6003", "input"]
        seletor_loja = ["wa-text-input#COMP6005", "input"]
        seletor_produto = ["wa-text-input#COMP6009", "input"]
        seletor_ultima_nf = ["wa-text-input#COMP6015", "input"]

        # ---> Seletores para as CÉLULAS da tabela
        seletor_celula_quantidade = ["#COMP7505", "input"]
        seletor_celula_dt_previsao = ["wa-text-input#COMP7507", "input"]
        seletor_celula_release = ["wa-text-input#COMP7510", "input"]
        seletor_celula_salvar = ["wa-button#COMP7516", "button"]
        seletor_botao_salvar = ["wa-button#COMP6022", "button"]
        seletor_botao_cancelar = ["wa-button#COMP6023", "button"]

        seletor_aviso_ultima_nf = ["wa-button#COMP7512", "button"]

        # ---> Espera inicial para o formulário de inclusão carregar
        try:
            print("Aguardando formulário de inclusão carregar...")
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, seletor_cod_cliente[0])
                )
            )
            print("Formulário pronto.")
        except TimeoutException:
            print(
                "ERRO CRÍTICO: O formulário de inclusão de provisões não carregou a tempo."
            )
            return [], list(df["C6_NUM"].dropna().astype(int).astype(str).unique()), []

        # ---> Laço for para tratamento e inclusão de cada linha da planilha
        for index, row in df.iterrows():
            self._verificar_parada()
            pedido_atual = None
            pn_voss = None
            try:
                if pd.isna(row["C6_CLI"]):
                    continue

                # ---> Mapeamento de cada coluna da planilha e suas respectivas informações por linha - Tratamento de cada campo com preenchimento com a quantidade ideal de zeros
                pedido_atual = str(int(row["C6_NUM"])).zfill(6)
                pn_voss = str(row["C6_PRODUTO"])
                cliente_cod = str(int(row["C6_CLI"])).zfill(6)
                loja = str(int(row["C6_LOJA"])).zfill(2)
                quantidade = str(int(row["C6_QTDVEN"])).zfill(9)

                # ---> Try: Verificação se há informação de data na linha de previsão de vendas
                try:
                    data_excel = row["C6_ENTREG"]
                    if pd.isna(data_excel):
                        print(
                            f"AVISO na linha {index + 2}: Data está em branco. Pulando."
                        )
                        if pedido_atual not in pedidos_falha:
                            pedidos_falha.add(pedido_atual)
                        if pn_voss not in pn_voss_falha:
                            pn_voss_falha.add(pn_voss)
                        continue

                    # ---> Tratamento da data de previsão de vendas
                    data_str = str(int(data_excel))
                    data_obj = datetime.strptime(data_str, "%Y%m%d")
                    data_obj = data_obj - timedelta(days=7)
                    data_date = data_obj.date()
                    data_final_string = data_date.strftime("%d/%m/%Y")
                    data_numeria_enviar = data_final_string.replace("/", "")
                    print(
                        f" ---> Data original: {data_str} | Data final formatada: {data_final_string}"
                    )

                except (ValueError, TypeError) as date_error:
                    print(
                        f"AVISO na linha {index + 2}: A data '{row['C6_ENTREG']}' é inválida. Pulando. Erro: {date_error}"
                    )
                    if pedido_atual not in pedidos_falha:
                        pedidos_falha.add(pedido_atual)
                    if pn_voss not in pn_voss_falha:
                        pn_voss_falha.add(pn_voss)
                    continue

                print(f"---> Processando linha {index + 2} | Produto: {pn_voss} ---")

                # ---> Preenchimento dos campos de cabeçalho
                # ---> Preenchimento código cliente
                time.sleep(0.5)
                campo_cliente_el = self._find_first_element(seletor_cod_cliente)
                campo_cliente_el.clear()
                campo_cliente_el.send_keys(cliente_cod)
                valor_atual_cliente = campo_cliente_el.get_attribute('value')
                if valor_atual_cliente == '' or valor_atual_cliente != cliente_cod:
                    campo_cliente_el.clear()
                    campo_cliente_el.send_keys(cliente_cod)
                print(f" ---> Código do Cliente {cliente_cod} preenchido")
                time.sleep(1)

                # ---> Preenchimento loja
                campo_loja_el = self._find_first_element(seletor_loja)
                campo_loja_el.clear()
                campo_loja_el.send_keys(loja)
                valor_atual_loja = campo_loja_el.get_attribute('value')
                if valor_atual_loja == '' or valor_atual_loja != loja:
                    campo_loja_el.clear()
                    campo_loja_el.send_keys(loja)
                print(f" ---> Loja {loja} preenchida.")
                time.sleep(1)

                # ---> Preenchimento código Voss
                campo_produto_el = self._find_first_element(seletor_produto)
                print(f"Quantidade de caracteres PN Voss: {len(pn_voss)}")
                if len(pn_voss) <= 15:
                    campo_produto_el.clear()
                    campo_produto_el.send_keys(pn_voss)
                    valor_atual_produto = campo_produto_el.get_attribute('value')
                    if valor_atual_produto == '' or valor_atual_produto != pn_voss:
                        campo_produto_el.clear()
                        campo_produto_el.send_keys(pn_voss)
                    campo_produto_el.send_keys(Keys.ENTER)
                    time.sleep(1)
                else:
                    campo_produto_el.send_keys(pn_voss)
                    if valor_atual_produto == '' or valor_atual_produto != pn_voss:
                        campo_produto_el.clear()
                        campo_produto_el.send_keys(pn_voss)   
                print(f" ---> Produto {pn_voss} preenchido e validado.")
                time.sleep(1)

                # ---> Identificação do campo ultima NF
                campo_ultima_nf = self._find_first_element(seletor_ultima_nf)
                campo_ultima_nf.click()
                valor_atual = campo_ultima_nf.get_attribute('value')
                valor_sem_espacos = valor_atual.replace(' ', '') 
                if valor_sem_espacos == '0':
                    campo_ultima_nf.clear() 
                    campo_ultima_nf.send_keys('1'.zfill(6))

                # campo_ultima_nf.click()
                time.sleep(0.5)
                campo_ultima_nf.send_keys(Keys.ENTER)
                time.sleep(0.5)
                self.actions.send_keys(Keys.ENTER).perform()
                time.sleep(1)

                # ---> Preenchimento da quantidade de itens
                campo_quantidade_el = self._find_first_element(
                    seletor_celula_quantidade
                )
                campo_quantidade_el.click()
                time.sleep(1)
                campo_quantidade_el.clear()
                time.sleep(1)
                campo_quantidade_el.send_keys(quantidade)
                valor_atual_quantidade = campo_quantidade_el.get_attribute('value') 
                if valor_atual_quantidade == '' or valor_atual_quantidade != quantidade:
                    campo_quantidade_el.click()
                    time.sleep(1)
                    campo_quantidade_el.clear()
                    time.sleep(1)
                    campo_quantidade_el.send_keys(quantidade)
                print(f"---> Quantidade: {quantidade} preenchida.")
                time.sleep(1)

                # ---> Preenchimento data de previsão de vendas
                campo_dt_previsao_el = self._find_first_element(
                    seletor_celula_dt_previsao
                )
                campo_dt_previsao_el.clear()
                time.sleep(1)
                campo_dt_previsao_el.send_keys(data_numeria_enviar)
                time.sleep(1)
                valor_atual_previsao = campo_dt_previsao_el.get_attribute('value') 
                if valor_atual_previsao == '' or valor_atual_previsao != data_numeria_enviar:
                    campo_dt_previsao_el.clear()
                    time.sleep(1)
                    campo_dt_previsao_el.send_keys(data_numeria_enviar)
                    time.sleep(1)
                print(f" ---> Data previsão {data_final_string} preenchida.")
                time.sleep(1)

                # ---> Preenchimento número release do pedido
                campo_release_el = self._find_first_element(seletor_celula_release)
                campo_release_el.clear()
                time.sleep(1)
                campo_release_el.send_keys(pedido_atual)
                valor_atual_pedido = campo_release_el.get_attribute('value')
                if valor_atual_pedido == '' or valor_atual_pedido != pedido_atual:
                    campo_release_el.clear()
                    time.sleep(1)
                    campo_release_el.send_keys(pedido_atual)
                print(f" ---> Número release {pedido_atual} preenchido.")
                time.sleep(1)

                # ---> Clicar no primeiro botão salvar
                campo_salvar_el = self._find_first_element(seletor_celula_salvar)
                campo_salvar_el.click()
                print(" ---> Clicando no botão salvar da tela de inclusão de previsão.")
                time.sleep(1)

                # ---> Clicar no segundo botão salvar
                self._find_first_element(seletor_botao_salvar).click()
                print(" ---> Clicando no botão salvar da rotina previsão de vendas.")
                time.sleep(2)

                # --- > Clicar no botão 'OK'após a inclusão
                self.actions.send_keys(Keys.ENTER).perform()
                print(" ---> Pop-up de sucesso confirmado.")
                time.sleep(1)

                # ---> Retorna a clicar no botão de incluir
                self.wait.until(
                    EC.text_to_be_present_in_element_value(
                        (By.CSS_SELECTOR, seletor_cod_cliente[0]), ""
                    )
                )

                print(f"Linha {index + 2} salva com sucesso para o produto {pn_voss}")
                if pedido_atual not in pedidos_sucesso:
                    pedidos_sucesso.append(pedido_atual)
                pedidos_falha.discard(pedido_atual)
                pn_voss_falha.discard(pn_voss)

            except Exception as e:
                print(f"ERRO ao processar a linha {index + 2}: {e}")
                if pedido_atual not in pedidos_falha:
                    pedidos_falha.add(pedido_atual)
                if pn_voss not in pedidos_falha:
                    pn_voss_falha.add(pn_voss)

                try:
                    print(" ---> Tentando recuperar do erro...")
                    self._find_first_element(seletor_botao_cancelar).click()
                    self.wait.until(
                        EC.invisibility_of_element_located((By.ID, "COMP6000"))
                    )
                    print(" ---> Diálogo de inclusão fechado.")

                    path_seletor_botao_incluir_browse = ["wa-button#COMP4566", "button"]
                    self._find_first_element(path_seletor_botao_incluir_browse).click()

                    self.wait.until(EC.presence_of_element_located((By.ID, "COMP6000")))
                    print(
                        " ---> Recuperação bem-sucedida. Continuando para a próxima linha."
                    )
                except Exception as cancel_error:
                    print(f" ---> ERRO CRÍTICO NA RECUPERAÇÃO. A automação será interrompida. Erro: {cancel_error}")
                    # Adiciona todos os itens restantes à lista de falhas
                    pedidos_restantes = df.loc[index:, "C6_NUM"].dropna().astype(str).unique()
                    pn_voss_restantes = df.loc[index:, "C6_PRODUTO"].dropna().astype(str).unique()
                    pedidos_falha.update(pedidos_restantes)
                    pn_voss_falha.update(pn_voss_restantes)
                    break 
                continue

        print(" ---> Processamento das previsões finalizado.")
        return pedidos_sucesso, list(pedidos_falha), list(pn_voss_falha)

    # ===> Função: Criação do email para usuário de acordo com o cadastro da previsão
    def enviar_email(self, pedidos_sucesso, pedidos_falha, pn_voss_falha, err=0, inf=None):
        # ---> Se erro fatal, envia e-mail de erro e sai
        if err == 1:
            print("Preparando para enviar email de notificação de ERRO FATAL.")
            try:
                outlook = Dispatch("outlook.application")
                mail = outlook.CreateItem(0)
                mail.To = ""
                mail.Subject = "ERRO FATAL - Ajuste Previsão de Vendas"
                mail.Attachments.Add(self.arquivo)
                html_body = f"""<p>Olá,</p>
                <p>A rotina de cadastro de previsão de vendas falhou com um erro crítico:</p><p><b>{inf}</b></p>
                <p>Att,
                <br>Bot VOSS</p>
                <br>VOSS Automotive Ltda.
                <br>Website: www.voss.net </br>"""
                mail.HTMLBody = html_body
                mail.Send()
                print("Email de erro fatal enviado!")
            except Exception as e:
                print(f"Ocorreu um erro ao tentar enviar o e-mail de erro fatal: {e}")
            return

        # ---> Se não houve itens válidos para processar na planilha
        if not pedidos_sucesso and not pedidos_falha:
            print(
                "Nenhum item válido encontrado na planilha para processar. Nenhum e-mail de status será enviado."
            )
            return

        print("Preparando para enviar email de notificação de status.")
        try:
            outlook = Dispatch("outlook.application")
            mail = outlook.CreateItem(0)
            mail.To = "juan.frois@voss.net"
            html_body = ""

            # ---> Caso 1: Todos os pedidos falharam
            if not pedidos_sucesso and pedidos_falha:
                pedidos_falha_str = ", ".join(map(str, sorted(pedidos_falha)))
                pn_voss_falha_str = ", ".join(map(str, sorted(pn_voss_falha)))
                mail.Subject = "FALHA - Ajuste Previsão de Vendas"
                html_body = f"""
                    <p>Olá,</p>
                    <p>A automação foi executada, mas <b>nenhuma previsão foi cadastrada com sucesso</b>.</p>
                    <p>Os seguintes pedidos continham itens que não puderam ser processados:<br>Pedido: {pedidos_falha_str} | PN Voss: {pn_voss_falha_str}</b>.</p>
                    <p>Verifique os logs no terminal para mais detalhes.</p>
                    <p>Att,<br>Sales Bot</p><br>VOSS Automotive Ltda.
                    <br>Website: www.voss.net </br>"""
                messagebox.showwarning(
                    "Atenção",
                    "Nenhuma previsão foi cadastrada com sucesso. Email de falha enviado.",
                )

            # ---> Caso 2: Sucesso no cadastro: total ou parcial
            elif pedidos_sucesso:
                pedidos_sucesso_str = ", ".join(map(str, sorted(pedidos_sucesso)))
                mail.Subject = (
                    f"Ajuste Previsão de Vendas - Pedido(s): {pedidos_sucesso_str}"
                )
                # CORREÇÃO: Tag <b> consertada
                corpo_email = f"""
                    <p>Olá,</p>
                    <p>As previsões para os pedidos <b>{pedidos_sucesso_str}</b> foram cadastradas com sucesso.</p>
                """
                if pedidos_falha:
                    pedidos_falha_str = ", ".join(map(str, sorted(pedidos_falha)))
                    pn_voss_falha_str = ", ".join(map(str, sorted(pn_voss_falha)))
                    corpo_email += f"<p><br>Atenção:</br> Os itens não puderam ser processados:<br>Pedido: {pedidos_falha_str} |<br>PN Voss: {pn_voss_falha_str}</br>.</p>"

                corpo_email += "<p>Att,<br>Sales Bot</p><br>VOSS Automotive Ltda.<br>Website: www.voss.net </br>"
                html_body = corpo_email
                messagebox.showinfo(
                    "Sucesso!", "Previsão de vendas processada e email enviado!"
                )
            if html_body: # Apenas envia se houver corpo de email
                mail.Attachments.Add(self.arquivo)
                mail.HTMLBody = html_body
                mail.Send()
                print("Email de status enviado com sucesso!")

        except Exception as e:
            print(f"Ocorreu um erro ao tentar enviar o e-mail: {e}")
            messagebox.showerror(
                "Erro de Email", f"Não foi possível enviar o e-mail: {e}"
            )
 
    # ===> Função: Execução de toda a automação, onde puxamos as funções acima criadas e organização de lógica de trabalho
    def executar(self):
        df = None
        tempo_maximo_espera = 15  # segundos
        tempo_atual = datetime.now()
        try:
            # ---> ETAPA: INICIALIZAÇÃO
            self._iniciar_driver()
            self._verificar_parada()
            print("\n ---> ACESSANDO PÁGINA DO MICROSIGA ---")
            self.driver.get("")
            self._verificar_parada()

            # ---> ETAPA 1: CLIQUE NO BOTÃO 'OK' INICIAL
            try:
                self._verificar_parada()
                path_botao_ok = [
                    "body > wa-dialog.startParameters.style-plastique",
                    "footer > wa-button:nth-child(2)",
                    "button",
                ]
                botao_ok = self._find_first_element(path_botao_ok)
                if botao_ok:
                    botao_ok.click()
                    print("Botão 'OK' inicial clicado.")
            except Exception as e:
                print(
                    f" ---> Aviso: Pop-up inicial não encontrado. Continuando... (Erro: {e})"
                )

            # ---> ETAPA 2: TRATAR POP-UP DE RESOLUÇÃO - SE EXISTIR
            while (datetime.now() - tempo_atual).total_seconds() < tempo_maximo_espera:
                self._verificar_parada()
                seletor_botao_fechar_resolucao = (By.CSS_SELECTOR, "wa-button#COMP3012")
                print(" ---> Verificando pop-up de aviso de resolução...")
                try:
                    botoes_fechar = self.short_wait.until(
                        EC.element_to_be_clickable((seletor_botao_fechar_resolucao))
                    )
                    if botoes_fechar:
                        print(
                            f" ---> Pop-up de resolução encontrado. Clicando em 'Fechar'."
                        )
                        botoes_fechar.click()
                        print(" ---> Pop-up de resolução fechado com sucesso.")
                        break
                    else:
                        print(
                            " ---> Pop-up de resolução não encontrado no tempo limite, continuando..."
                        )
                except TimeoutException:
                    time.sleep(1)
                    continue
                except Exception as e:
                    print(
                        f"\n ---> Aviso: Ocorreu um erro ao tentar fechar o pop-up de resolução. Erro: {e}"
                    )

            # --- ETAPA 3 e 4: TELA DE LOGIN ---
            try:
                time.sleep(1)
                self._verificar_parada()
                print("\n ---> INICIANDO TELA DE LOGIN ---")
                login_host = self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "wa-webview[id^='COMP']")
                    )
                )
                login_shadow_root = self.driver.execute_script(
                    "return arguments[0].shadowRoot", login_host
                )
                iframe = login_shadow_root.find_element(By.CSS_SELECTOR, "iframe")
                self.driver.switch_to.frame(iframe)
                print("iFrame de Login selecionado.")
                self.wait.until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "input[name='login']")
                    )
                ).send_keys(self.login)
                self.driver.find_element(
                    By.CSS_SELECTOR, "input[name='password']"
                ).send_keys(self.senha)
                self.driver.find_element(By.CSS_SELECTOR, ".po-button").click()
                print("Login realizado com sucesso.")
            finally:
                self.driver.switch_to.default_content()
                print("Contexto do driver restaurado para a página principal.")

            # --- ETAPA 5: TELA DE AMBIENTE ---
            try:
                time.sleep(3)
                self._verificar_parada()
                css_selector_ambiente_host = "wa-webview[src*='preindex_env_voss']"
                print("\n--- INICIANDO ETAPA DE CONFIGURAÇÃO DE AMBIENTE E SESSÃO ---")
                ambiente_host = self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, css_selector_ambiente_host)
                    )
                )
                ambiente_shadow_root = self.driver.execute_script(
                    "return arguments[0].shadowRoot", ambiente_host
                )
                ambiente_iframe = ambiente_shadow_root.find_element(
                    By.CSS_SELECTOR, "iframe"
                )
                self.driver.switch_to.frame(ambiente_iframe)
                print("Foco no Iframe de Ambiente.")

                # 1. Preencher campo Ambiente
                css_selector_todos_os_campos_lookup = "input[id^='po-lookup']"
                todos_os_campos_lookup = self.wait.until(
                    EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, css_selector_todos_os_campos_lookup)
                    )
                )
                if len(todos_os_campos_lookup) < 3:
                    raise Exception(
                        "Campo 'Ambiente' (terceiro lookup) não encontrado."
                    )
                campo_ambiente = todos_os_campos_lookup[2]
                if campo_ambiente.get_attribute("value") != "10":
                    campo_ambiente.clear()
                    campo_ambiente.send_keys("10")
                    print("Ambiente '10' inserido.")
                else:
                    print("Ambiente '10' já estava preenchido.")
                campo_ambiente.send_keys(Keys.TAB)

                # 2. Ativar o segundo Switch
                css_selector_todos_os_switches = "div[role='switch']"
                todos_os_switches = self.wait.until(
                    EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, css_selector_todos_os_switches)
                    )
                )
                if len(todos_os_switches) < 2:
                    raise Exception(
                        "Segundo switch (Manter Informações) não encontrado."
                    )
                campo_sessao = todos_os_switches[1]
                if campo_sessao.get_attribute("aria-checked") == "false":
                    campo_sessao.click()
                    print("Segundo switch (Manter Informações) ativado.")
                else:
                    print("Segundo switch já estava ativado.")

                # 3. Clicar no terceiro botão "Entrar"
                css_selector_todos_os_botoes = "button.po-button"
                todos_os_botoes = self.wait.until(
                    EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, css_selector_todos_os_botoes)
                    )
                )
                print(f"Foram identificados:{len(todos_os_botoes)}")
                if len(todos_os_botoes) < 3:
                    raise Exception("Terceiro botão (Enter) não foi encontrado.")
                terceiro_botao_entrar = todos_os_botoes[2]
                self.driver.execute_script(
                    "arguments[0].click();", terceiro_botao_entrar
                )
                print("Terceiro botão(Entrar) clicado.")
            finally:
                self.driver.switch_to.default_content()
                print("Login finalizado. Contexto restaurado para a página principal.")

            # --- ETAPA 6: NAVEGAÇÃO PARA A ROTINA ---
            self._verificar_parada()
            self.navegar_rotina(modulo="MATASC4")

            # Chama o tratamento de pop-ups APÓS a navegação.
            self._verificar_parada()
            self.tratar_pop_ups()

            # --- ETAPA 7: INTERAÇÃO COM A ROTINA ---
            self._verificar_parada()
            print("\n ---> CLICANDO EM 'INCLUIR' NA ROTINA: MATASC4 ---")
            path_seletor_botao_incluir = ["wa-button#COMP4566", "button"]

            # Espera o botão Incluir da tela de browse estar clicável
            self.wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, path_seletor_botao_incluir[0])
                )
            ).click()

            # --- ETAPA 8: PROCESSAMENTO DO ARQUIVO ---
            self._verificar_parada()
            df = self.ler_arquivo()
            pedidos_sucesso, pedidos_falha, pn_voss_falha = self.processar_provisoes(df)
            self.enviar_email(pedidos_sucesso, pedidos_falha, pn_voss_falha)

            return {
                "sucesso": True,
                "mensagem": "Automação concluída com sucesso!",
            }

        except InterruptedError as e:
            print(f"AVISO: {e}")
            return {"sucesso": False, "mensagem": str(e)}
        except Exception as e:
            print(f"ERRO FATAL NA EXECUÇÃO: {e}")
            import traceback

            traceback.print_exc()
            msg_erro = f"Exceção: {type(e).__name__}\nDetalhes: {str(e)}\nTraceback: {traceback.format_exc()}"
            self.enviar_email([], [],[], err=1, inf=msg_erro)
            return {"sucesso": False, "mensagem": f"Erro fatal: {e}"}

        finally:
            if self.driver and not self.evento_parar.is_set():
                self.driver.quit()
