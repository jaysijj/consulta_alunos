import tkinter as tk
from tkinter import ttk
import ttkbootstrap as tb
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from tkinter import filedialog
import os
import threading

class SeleniumApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Consulta Alunos")
        self.root.geometry("580x750")
        
        # Configurar estilo
        self.style = tb.Style(theme="darkly")
        
        # Frame principal
        self.main_frame = tb.Frame(root)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        # Título
        self.title_label = tb.Label(
            self.main_frame, 
            text="Consulta de Alunos", 
            font=('Helvetica', 18, 'bold')
        )
        self.title_label.pack(pady=10)
        
        # Frame de credenciais
        self.credenciais_frame = tb.LabelFrame(
            self.main_frame, 
            text="Credenciaiss", 
            bootstyle="info"
        )
        self.credenciais_frame.pack(fill="x", pady=10, padx=5)

        # Configuração das colunas para expansão
        self.credenciais_frame.grid_columnconfigure(1, weight=1)  # Expande a coluna do usuário
        self.credenciais_frame.grid_columnconfigure(3, weight=1)  # Expande a coluna da senha
        
        # Campos de usuário e senha
        tb.Label(self.credenciais_frame, text="Usuário:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.usuario_entry = tb.Entry(self.credenciais_frame)
        self.usuario_entry.grid(row=0, column=1, padx=(5, 15), pady=5, sticky="ew")
        self.usuario_entry.insert(0, os.getenv("SI3_USER", ""))
        
        tb.Label(self.credenciais_frame, text="Senha:").grid(row=0, column=2, padx=(15, 5), pady=5, sticky="e")
        self.senha_entry = tb.Entry(self.credenciais_frame, show="*")
        self.senha_entry.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        self.senha_entry.insert(0, os.getenv("SI3_PASSWORD", ""))
        
        # Frame de parâmetros de busca
        self.busca_frame = tb.LabelFrame(
            self.main_frame, 
            text="Parâmetros de Busca", 
            bootstyle="info"
        )
        self.busca_frame.pack(fill="x", pady=10, padx=5)

        # Configuração das colunas para expansão
        self.busca_frame.grid_columnconfigure(1, weight=1)
        self.busca_frame.grid_columnconfigure(3, weight=1)

        # Ano de ingresso (Par 1 - Campo 1)
        tb.Label(self.busca_frame, text="Ano de Ingresso:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.ano_ingresso_entry = tb.Entry(self.busca_frame)
        self.ano_ingresso_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.ano_ingresso_entry.insert(0, "")

        # Centro (Par 1 - Campo 2)
        tb.Label(self.busca_frame, text="Centro:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.centro_options = [
            ("Campus", "10411"),
        ]
        self.centro_combobox = tb.Combobox(
            self.busca_frame, 
            values=[opt[0] for opt in self.centro_options],
            state="readonly"
        )
        self.centro_combobox.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        self.centro_combobox.current(0)

        # Status (Par 2 - Campo 1)
        tb.Label(self.busca_frame, text="Status:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.status_options = [
            ("ATIVO", "1"), ("CADASTRADO", "2"), ("CANCELADO", "6"), ("CONCLUÍDO", "3"), ("DEFENDIDO", "12"),
            ("EM HOMOLOGAÇÃO", "11"), ("EM MATRÍCULA INSTITUCIONAL", "23"), ("EM MOBILIDADE ACADÊMICA","22"),
            ("EXCLUÍDO", "10"), ("FORMANDO", "8"), ("GRADUANDO", "9"), ("MUDANÇA DE NÍVEL", "21"), ("TRANCADO", "5")
        ]
        self.status_combobox = tb.Combobox(
            self.busca_frame, 
            values=[opt[0] for opt in self.status_options],
            state="readonly"
        )
        self.status_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.status_combobox.current(0)

        # Curso (Par 2 - Campo 2)
        tb.Label(self.busca_frame, text="Curso:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.curso_options = [
            ("CIÊNCIA DA COMPUTAÇÃO", "1878933"),
            ("ENGENHARIA AMBIENTAL", "2097265"),
            ("ENGENHARIA CIVIL", "2097286"),
            ("ENGENHARIA DE MINAS", "2118717"),
            ("SISTEMAS DE INFORMAÇÃO", "2097205")
        ]
        self.curso_combobox = tb.Combobox(
            self.busca_frame, 
            values=[opt[0] for opt in self.curso_options],
            state="readonly"
        )
        self.curso_combobox.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
        self.curso_combobox.current(2)

        # Frame para fechar avisos/popups
        self.popup_frame = tb.LabelFrame(
            self.main_frame, 
            text="Retirar Popups/Avisos novos", 
            bootstyle="info"
        )
        self.popup_frame.pack(fill="x", pady=10, padx=5)
        
        # XPATH para fechar popup 1
        tb.Label(self.popup_frame, text="XPATH Popup 1:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.popup1_entry = tb.Entry(self.popup_frame)
        self.popup1_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.popup1_entry.insert(0, "")

        # XPATH para fechar popup 2
        tb.Label(self.popup_frame, text="XPATH Popup 2:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.popup2_entry = tb.Entry(self.popup_frame)
        self.popup2_entry.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        self.popup2_entry.insert(0, "")

        # XPATH para fechar popup 3
        tb.Label(self.popup_frame, text="XPATH Popup 3:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.popup3_entry = tb.Entry(self.popup_frame)
        self.popup3_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.popup3_entry.insert(0, "")

        # XPATH para fechar popup 4
        tb.Label(self.popup_frame, text="XPATH Popup 4:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.popup4_entry = tb.Entry(self.popup_frame)
        self.popup4_entry.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
        self.popup4_entry.insert(0, "")

        # Configurar pesos das colunas para melhor distribuição
        self.popup_frame.grid_columnconfigure(1, weight=1)
        self.popup_frame.grid_columnconfigure(3, weight=1)

        # Diretório de Download
        self.download_frame = tb.LabelFrame(
            self.main_frame, 
            text="Diretório de Download", 
            bootstyle="info"
        )
        self.download_frame.pack(fill="x", pady=10, padx=5)
        
        # Campo para exibir o diretório selecionado
        self.download_dir = tb.Entry(self.download_frame)
        self.download_dir.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        
        # Botão para selecionar diretório
        self.select_dir_btn = tb.Button(
            self.download_frame,
            text="Selecionar",
            command=self.selecionar_diretorio,
            bootstyle="primary"
        )
        self.select_dir_btn.pack(side="right", padx=5, pady=5)

        # Botão de execução
        self.executar_button = tb.Button(
            self.main_frame,
            text="Executar Consulta",
            command=self.executar_consulta,
            bootstyle="success"
        )
        self.executar_button.pack(pady=20)
        
        # Área de logs
        self.log_frame = tb.LabelFrame(
            self.main_frame, 
            text="Logs de Execução", 
            bootstyle="info"
        )
        self.log_frame.pack(fill="both", expand=True, pady=10, padx=5)
        
        self.log_text = tk.Text(self.log_frame, height=10)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Adicionar scrollbar
        scrollbar = tb.Scrollbar(self.log_text)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_text.yview)
        
    def log(self, message):
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.root.update()

    def fechar_popups(self, driver):
        """Tenta fechar popups/avisos usando os XPATHs configurados"""
        for xpath_entry in [self.popup1_entry, self.popup2_entry, self.popup3_entry, self.popup4_entry]:
            xpath = xpath_entry.get().strip()
            if xpath:
                try:
                    elements = driver.find_elements(By.XPATH, xpath)
                    if elements:
                        for element in elements:
                            try:
                                element.click()
                                self.log("Popup fechado com sucesso usando XPATH")
                            except:
                                # Se não puder clicar, tenta executar JavaScript
                                driver.execute_script("arguments[0].click();", element)
                                self.log("Popup fechado usando JavaScript")
                except Exception as e:
                    self.log(f"Erro ao tentar fechar popup: {str(e)}")

    def selecionar_diretorio(self):
        diretorio = filedialog.askdirectory()
        if diretorio:
            # Converter para caminho absoluto e normalizar para o sistema
            diretorio = os.path.abspath(diretorio)
            self.download_dir.delete(0, "end")
            self.download_dir.insert(0, diretorio)

    def _executar_consulta_thread(self):
        """Método que será executado em uma thread separada"""
        try:
            # Chamar a função original com os parâmetros necessários
            self._executar_consulta_core()
        except Exception as e:
            self.log(f"Erro na thread: {str(e)}")
        finally:
            # Reabilitar o botão quando terminar
            self.root.after(0, lambda: self.executar_button.config(state="normal"))

    def executar_consulta(self):
        self.executar_button.config(state="disabled")

        thread = threading.Thread(target=self._executar_consulta_thread, daemon=True)
        thread.start()

    def _executar_consulta_core(self):
        usuario = self.usuario_entry.get()
        senha = self.senha_entry.get()
        ano_ingresso = self.ano_ingresso_entry.get()
        
        # Obter valores selecionados
        status_index = self.status_combobox.current()
        centro_index = self.centro_combobox.current()
        curso_index = self.curso_combobox.current()
        
        status_value = self.status_options[status_index][1]
        centro_value = self.centro_options[centro_index][1]
        curso_value = self.curso_options[curso_index][1]

        # Obter diretório de download
        download_path = self.download_dir.get()
        if not download_path:
            self.log("Por favor, selecione um diretório para download")
            return
            
        # Configurar opções de download do Chrome
        chrome_options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": download_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--safebrowsing-disable-download-protection")

        self.log("Iniciando automação...")
        
        try:
            # Inicializando o WebDriver
            self.log("Configurando o navegador...")
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
            wait = WebDriverWait(driver, 10)

            # 1. Abrir a página inicial
            self.log("Acessando o SI3...")
            driver.get(os.getenv("SI3_SITE"))
            time.sleep(2)
            
            # Preencher credenciais
            self.log("Preenchendo credenciais...")
            campo_usuario = wait.until(EC.element_to_be_clickable((By.NAME, 'user.login')))
            campo_usuario.click()
            campo_usuario.send_keys(usuario)
            
            campo_senha = wait.until(EC.element_to_be_clickable((By.NAME, 'user.senha')))
            campo_senha.click()
            campo_senha.send_keys(senha)
            
            botao_entrar = wait.until(EC.element_to_be_clickable((By.NAME, 'entrar')))
            botao_entrar.click()

            self.fechar_popups(driver)
            
            # Navegação no sistema
            self.log("Navegando para o portal de coordenação...")
            selecionar_campo_createus = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="j_id_jsp_1572861391_1"]/section/ul/li[2]/a')))
            selecionar_campo_createus.click()
            
            portal_coord_graduacao = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'portal_graduacao')))
            portal_coord_graduacao.click()
            
            # Selecionar curso
            self.log(f"Selecionando curso: {self.curso_options[curso_index][0]}")
            elemento_select = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//select[contains(@name, 'j_id_jsp')]")
            ))
            selecionar_curso = Select(elemento_select)
            selecionar_curso.select_by_value(curso_value)
            
            # Acessar consulta avançada
            self.log("Acessando consulta avançada...")
            aba_aluno = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="cmAction-39"]/span[1]')))
            aba_aluno.click()
            
            consulta_avancada = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="cmAction-40"]/td[2]')))
            consulta_avancada.click()
            
            # Preencher parâmetros de busca
            if (ano_ingresso):
                self.log("Preenchendo parâmetros de busca...")
                ano_ingresso_field = wait.until(EC.element_to_be_clickable((By.ID, 'busca:inputAnoIngresso')))
                ano_ingresso_field.click()
                ano_ingresso_field.send_keys(ano_ingresso)
            
            status_field = wait.until(EC.element_to_be_clickable((By.ID, 'busca:selectStatus')))
            selecionar_status = Select(status_field)
            selecionar_status.select_by_value(status_value)
            
            centro_field = wait.until(EC.element_to_be_clickable((By.ID, 'busca:selectCentro')))
            selecionar_centro = Select(centro_field)
            selecionar_centro.select_by_value(centro_value)
            
            curso_field = wait.until(EC.element_to_be_clickable((By.ID, 'busca:selectCurso')))
            selecionar_curso_consulta = Select(curso_field)
            selecionar_curso_consulta.select_by_value(curso_value)
            
            # Executar busca
            self.log("Executando busca...")
            buscar = wait.until(EC.element_to_be_clickable((By.NAME, 'busca:j_id_jsp_1556043697_440')))
            buscar.click()

            # Verificar se há mensagem de erro
            try:
                # Verificar rapidamente se a mensagem de erro está presente
                error_message = driver.find_elements(By.XPATH, '//ul[@class="warning"]/li')
                if error_message:
                    self.log(f"ERRO: {error_message[0].text}")
                    driver.quit()
                    return
            except:
                pass  # Se não encontrar o elemento de erro, continua normalmente
            
            # Processar resultados
            self.log("Processando resultados...")
            encontrar_numero_pdf = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="busca"]/table[2]/caption')))
            numero_pdf = encontrar_numero_pdf.text.split(":")[1].strip()
            numero_pdf_int = int(numero_pdf)
            
            self.log(f"Encontrados {numero_pdf_int} registros. Iniciando download dos PDFs...")
            
            for pdf in range(1, numero_pdf_int + 1):
                self.log(f"Baixando PDF {pdf} de {numero_pdf_int}...")
                try:
                    buscar = wait.until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="busca"]/table[2]/tbody/tr[{pdf}]/td[5]/a/img')))
                    buscar.click()
                    time.sleep(1)  # Pequena pausa entre downloads
                except Exception as e:
                    self.log(f"Erro ao baixar PDF {pdf-1} !!")
                    return
            self.log("Processo concluído com sucesso!")
            
        except Exception as e:
            self.log(f"Ocorreu um erro: {str(e)}")
        finally:
            try:
                driver.quit()
            except:
                pass

if __name__ == "__main__":
    root = tb.Window(themename="darkly")
    app = SeleniumApp(root)
    root.mainloop()