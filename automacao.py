import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import csv
import re
import os
import math
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
import time
from selenium.webdriver.chrome.options import Options  # Importar Options

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação Requisição - Selenium")
        self.root.geometry("1000x500")

        self.csv_path = None
        self.itens_pulados = []
        self.itens_ajustados = []  # Nova lista para itens ajustados
        self.itens_repetidos = []  # Lista para armazenar itens repetidos removidos

        self.btn_load_csv = tk.Button(root, text="Selecionar arquivo CSV", command=self.load_csv)
        self.btn_load_csv.pack(pady=10)

        self.lbl_csv = tk.Label(root, text="Nenhum arquivo selecionado")
        self.lbl_csv.pack()

        self.btn_run = tk.Button(root, text="Criar requisição", command=self.run_automation_thread, state=tk.DISABLED)
        self.btn_run.pack(pady=10)

        self.lbl_pulados = tk.Label(root, text="Relatório:")
        self.lbl_pulados.pack(pady=(20,0))

        self.txt_pulados = scrolledtext.ScrolledText(root, height=35, width=150)
        self.txt_pulados.pack(padx=10, pady=5)
        
        # Configurar tags para cores
        self.txt_pulados.tag_config("amarelo", background="#F8F68F")
        self.txt_pulados.tag_config("vermelho", background="#FFCCCC")  # Vermelho claro
        self.txt_pulados.tag_config("verde", background="#CCFFCC")    # Verde claro
        self.txt_pulados.tag_config("azul", background="#CCE5FF")     # Azul claro
        self.txt_pulados.tag_config("laranja", background="#FFE5CC")  # Laranja claro para repetidos
        self.txt_pulados.tag_config("titulo", font=("Arial", 10, "bold"))

    def load_csv(self):
        path = filedialog.askopenfilename(
            title="Selecione o arquivo CSV",
            filetypes=[("CSV files", "*.csv"), ("Todos os arquivos", "*.*")]
        )
        if path:
            self.csv_path = path
            nome_arquivo = os.path.basename(path)
            self.lbl_csv.config(text=f"Arquivo selecionado: {nome_arquivo}")
            self.btn_run.config(state=tk.NORMAL)
            self.txt_pulados.delete("1.0", tk.END)
            self.itens_pulados = []
            self.itens_ajustados = []
            self.itens_repetidos = []

    def run_automation_thread(self):
        self.btn_run.config(state=tk.DISABLED)
        self.txt_pulados.delete("1.0", tk.END)
        self.itens_pulados = []
        self.itens_ajustados = []
        self.itens_repetidos = []
        threading.Thread(target=self.run_automation, daemon=True).start()

    def buscar_item(self, driver, wait, codigo, max_tentativas=3):
        """Tenta buscar um item várias vezes antes de considerar como não encontrado"""
        for tentativa in range(max_tentativas):
            try:
                time.sleep(1.5) 
                codigo_input = wait.until(EC.presence_of_element_located((By.NAME, "CD_CODIGO")))
                codigo_input.clear()
                codigo_input.send_keys(codigo)

                localizar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn-green') and contains(text(), 'Localizar')]")))
                driver.execute_script("arguments[0].scrollIntoView(true);", localizar_btn)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", localizar_btn)

                # Aguarda o resultado da busca
                wait.until(
                    EC.any_of(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "div.datagrid-cell-content-wrapper")),
                        EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'datagrid-empty-msg') and contains(text(), 'Esta lista está vazia.')]"))
                    )
                )
                time.sleep(0.5)

                # Verifica se a lista está vazia
                empty_elements = driver.find_elements(By.XPATH, "//span[contains(@class, 'datagrid-empty-msg') and contains(text(), 'Esta lista está vazia.')]")
                if any(el.is_displayed() for el in empty_elements):
                    if tentativa < max_tentativas - 1:
                        time.sleep(1)  # Pequena pausa antes de tentar novamente
                        continue
                    else:
                        return False  # Item não encontrado após todas as tentativas
                
                # Se encontrou o item, retorna True
                return True

            except Exception as e:
                if tentativa < max_tentativas - 1:
                    time.sleep(1)
                    continue
                else:
                    return False
        
        return False

    def arredondar_para_multiplo_de_5(self, valor):
        """Arredonda um valor para cima para o próximo múltiplo de 5"""
        return math.ceil(valor / 5) * 5

    def run_automation(self):
        if not self.csv_path:
            messagebox.showerror("Erro", "Nenhum arquivo CSV selecionado.")
            self.btn_run.config(state=tk.NORMAL)
            return

        driver = None
        try:
            nome_base = os.path.splitext(os.path.basename(self.csv_path))[0]

            # Dicionários com os limites de quantidade por código
            LIMITES_QUANTIDADE = {
                # Limite de 1
                1: {
                    '468', '60', '144', '74', '144', '1165', '1114', '1564', '258', '259',
                    '260', '229', '753', '223', '158', '393', '1716', '1134', '147', '149',
                    '1503', '401', '403', '146', '99', '648', '532', '526', '105', '896',
                    '176', '754', '171', '407', '177', '241', '39', '1170', '1651', '245',
                    '246', '248', '249', '250', '251', '253', '254', '1474', '1175', '1149'
                },
                # Limite de 2
                2: {
                    '607', '606', '150', '1427', '319', '322', '323', '324', '325', '327',
                    '328', '1242', '1733', '228', '1118', '911', '997', '1722', '1723', '1724'
                },
                # Limite de 3
                3: {
                    '1094', '1203', '1292', '1293', '1294', '1295', '1305', '1521', '471', '472',
                    '473', '474', '476', '477', '479', '480', '481', '482', '836', '1561',
                    '1599', '175', '31', '278', '154', '101'
                },
                # Limite de 4
                4: {
                    '349'
                },
                # Limite de 5
                5: {
                    '1379', '619', '620', '621'
                },
                # Limite de 6
                6: {
                    '408'
                },
                # Limite de 10
                10: {
                    '244', '640', '1180', '108', '110', '1275', '137', '136', '1735','142'
                },
                # Limite de 20
                20: {
                    '452', '1106'
                }
            }
            
            # Códigos que devem ser arredondados para o próximo múltiplo de 5
            CODIGOS_ARREDONDAR_MULTIPLO_5 = {
                '104', '624', '625', '626', '421', '424','435', '1173', '642','419','1440'
                '1185', '1709', '336', '485', '492', '496', '499', '511', '517','420','746','9','579','580','575','563','505','507'
            }
            
            # Códigos com quantidade fixa
            CODIGOS_QUANTIDADE_FIXA = {
                '230': 20  # Código 230 sempre quantidade 20
            }
            
            # Cria um dicionário reverso para busca rápida
            LIMITE_POR_CODIGO = {}
            for limite, codigos in LIMITES_QUANTIDADE.items():
                for codigo in codigos:
                    LIMITE_POR_CODIGO[codigo] = limite
            
            # Dicionário para armazenar itens agrupados por código
            itens_agrupados = {}
            itens_repetidos_detalhes = []
            total_linhas_csv = 0
            
            with open(self.csv_path, newline='', encoding='latin1') as csvfile:
                # Primeiro vamos ler os cabeçalhos manualmente
                primeira_linha = csvfile.readline().strip()
                cabecalhos = primeira_linha.split(';')
                cabecalhos = [cab.strip() for cab in cabecalhos]
                
                # Volta ao início do arquivo
                csvfile.seek(0)
                reader = csv.DictReader(csvfile, delimiter=';', fieldnames=cabecalhos)
                next(reader)  # Pula a linha de cabeçalho
                
                # Encontra o índice das colunas importantes
                coluna_quebra = None
                coluna_descricao = None
                coluna_requisitado = None
                
                for i, cab in enumerate(cabecalhos):
                    if 'quebra' in cab.lower():
                        coluna_quebra = cab
                    elif 'descri' in cab.lower():
                        coluna_descricao = cab
                    elif 'requisitado' in cab.lower():
                        coluna_requisitado = cab
                
                # Usa os nomes padrão se não encontrou
                if not coluna_quebra:
                    coluna_quebra = 'Quebra'
                if not coluna_descricao:
                    coluna_descricao = 'Descriï¿½ï¿½o'
                if not coluna_requisitado:
                    coluna_requisitado = 'Requisitado'
                
                for row in reader:
                    total_linhas_csv += 1
                    
                    try:
                        codigo_raw = row[coluna_quebra].strip() if coluna_quebra in row else ''
                        descricao = row[coluna_descricao].strip() if coluna_descricao in row else 'Descrição não encontrada'
                        quantidade_str = row[coluna_requisitado].strip().replace(',', '.') if coluna_requisitado in row else ''

                        match = re.search(r'\d+', codigo_raw)
                        codigo = match.group(0) if match else ''

                        if codigo and quantidade_str:
                            try:
                                quantidade = float(quantidade_str)
                                
                                # APLICAÇÃO DAS REGRAS ESPECIAIS
                                quantidade_original = quantidade
                                regra_aplicada = "Nenhuma regra especial"
                                foi_ajustado = False
                                
                                # 1. Código 230: quantidade fixa de 20 (PRIMEIRO - prioridade máxima)
                                if codigo == '230':
                                    quantidade = 20.0
                                    regra_aplicada = "Quantidade fixa de 20"
                                    foi_ajustado = True
                                
                                # 2. Pilhas divide por 2 e arredonda para cima
                                elif codigo in ['262', '263']:
                                    quantidade = math.ceil(quantidade / 2)
                                    regra_aplicada = f"Dividido por 2 e arredondado"
                                    foi_ajustado = True
                                
                                # 3. Códigos que devem ser arredondados para múltiplo de 5
                                elif codigo in CODIGOS_ARREDONDAR_MULTIPLO_5:
                                    quantidade_arredondada = self.arredondar_para_multiplo_de_5(quantidade)
                                    if quantidade_arredondada != quantidade:
                                        quantidade = quantidade_arredondada
                                        regra_aplicada = f"Arredondado para múltiplo de 5"
                                        foi_ajustado = True
                                
                                # 4. Verifica se o código tem limite de quantidade
                                elif codigo in LIMITE_POR_CODIGO:
                                    limite = LIMITE_POR_CODIGO[codigo]
                                    if quantidade > limite:
                                        regra_aplicada = f"Limitado para {limite}"
                                        quantidade = float(limite)
                                        foi_ajustado = True
                                
                                # 5. Conversões normais para outros códigos especiais
                                elif codigo in ['1604', '1601', '1603', '490']:
                                    quantidade *= 50
                                    regra_aplicada = "Multiplicado por 50"
                                    foi_ajustado = True
                                elif codigo in ['890', '889', '15','1070','1071']:
                                    quantidade *= 100
                                    regra_aplicada = "Multiplicado por 100"
                                    foi_ajustado = True
                                
                                # Verifica se o código já existe no dicionário
                                if codigo in itens_agrupados:
                                    # Verifica se a quantidade atual é maior que a armazenada
                                    if quantidade > itens_agrupados[codigo]['quantidade']:
                                        # Armazena o item antigo na lista de repetidos
                                        item_antigo = itens_agrupados[codigo]
                                        itens_repetidos_detalhes.append({
                                            'Codigo': codigo,
                                            'Descricao': item_antigo['descricao'],
                                            'Quantidade_Antiga': f"{item_antigo['quantidade_original']:.2f}".replace('.', ','),
                                            'Quantidade_Nova': f"{quantidade_original:.2f}".replace('.', ','),
                                            'Motivo': 'Substituído por quantidade maior'
                                        })
                                        
                                        # Atualiza com o novo item (maior quantidade)
                                        itens_agrupados[codigo] = {
                                            'quantidade': quantidade,
                                            'quantidade_original': quantidade_original,
                                            'descricao': descricao,
                                            'regra_aplicada': regra_aplicada,
                                            'foi_ajustado': foi_ajustado
                                        }
                                    else:
                                        # A quantidade atual é menor ou igual, descarta este item
                                        itens_repetidos_detalhes.append({
                                            'Codigo': codigo,
                                            'Descricao': descricao,
                                            'Quantidade_Antiga': f"{quantidade_original:.2f}".replace('.', ','),
                                            'Quantidade_Atual_Maior': f"{itens_agrupados[codigo]['quantidade_original']:.2f}".replace('.', ','),
                                            'Motivo': 'Ignorado por quantidade menor'
                                        })
                                else:
                                    # Primeira ocorrência deste código
                                    itens_agrupados[codigo] = {
                                        'quantidade': quantidade,
                                        'quantidade_original': quantidade_original,
                                        'descricao': descricao,
                                        'regra_aplicada': regra_aplicada,
                                        'foi_ajustado': foi_ajustado
                                    }
                                
                            except ValueError:
                                motivo = "Quantidade inválida"
                                self.itens_pulados.append({'Codigo': codigo, 'Descricao': descricao, 'Quantidade': quantidade_str, 'Motivo': motivo})
                        else:
                            motivo = "Dados incompletos"
                            self.itens_pulados.append({'Codigo': codigo, 'Descricao': descricao, 'Quantidade': quantidade_str, 'Motivo': motivo})
                    
                    except KeyError as e:
                        continue

            # Prepara a lista final de itens para processamento
            itens_csv = []
            for codigo, dados in itens_agrupados.items():
                # Adiciona à lista de ajustados se foi modificado
                if dados['foi_ajustado']:
                    self.itens_ajustados.append({
                        'Codigo': codigo,
                        'Descricao': dados['descricao'],
                        'Quantidade_Original': f"{dados['quantidade_original']:.2f}".replace('.', ','),
                        'Quantidade_Ajustada': f"{dados['quantidade']:.2f}".replace('.', ','),
                        'Regra': dados['regra_aplicada']
                    })
                
                # Formata a quantidade final
                quantidade_formatada = f"{dados['quantidade']:.2f}".replace('.', ',')
                itens_csv.append((codigo, quantidade_formatada, dados['descricao']))
            
            # Armazena os itens repetidos removidos
            self.itens_repetidos = itens_repetidos_detalhes

            # Configurar opções do Chrome para zoom de 80%
            chrome_options = Options()
            chrome_options.add_argument("--force-device-scale-factor=0.8")
            chrome_options.add_argument("--high-dpi-support=0.8")
            chrome_options.add_argument("--window-size=1920,1080")  # Opcional: definir tamanho da janela
            
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 30)

            driver.get("https://zermatt.digisystem.cloud/#/")
            
            # Aplicar zoom via JavaScript também (para garantir)
            driver.execute_script("document.body.style.zoom='80%'")

            wait.until(EC.element_to_be_clickable((By.ID, "loginUsername"))).send_keys("lucas.godoy")
            time.sleep(0.3)
            wait.until(EC.element_to_be_clickable((By.ID, "loginPassword"))).send_keys("Tasy!007")
            time.sleep(0.3)
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input.btn-green.w-login-button.w-login-button--green"))).click()
            time.sleep(1)

            try:
                close_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.w-header-tab__close")))
                close_btn.click()
                time.sleep(1)
            except TimeoutException:
                pass

            wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Requisição de Materiais e Medicamentos']"))).click()
            time.sleep(1)

            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Adicionar') and contains(@class, 'btn-blue')]"))).click()
            time.sleep(1)

            wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'w-listbox__label')]//span[normalize-space()='Transferência - Saída']"))).click()
            time.sleep(0.3)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Consumo']"))).click()
            time.sleep(1)

            campo_observacao = wait.until(EC.element_to_be_clickable((By.NAME, "DS_OBSERVACAO")))
            campo_observacao.click()
            campo_observacao.clear()
            campo_observacao.send_keys(nome_base)
            time.sleep(1)

            wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'wbutton-container') and contains(@class, 'btn-blue')]"))).click()
            time.sleep(1) 

            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Cancelar']"))).click()
            time.sleep(1)

            try:
                wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'dialog_ok_button') and normalize-space()='Sim']"))).click()
            except:
                wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'rg-label') and normalize-space()='Sim']"))).click()
            time.sleep(1)

            botoes = driver.find_elements(By.XPATH, "//button[contains(text(),'Adicionar') and contains(@class,'btn-blue')]")
            for btn in botoes:
                if btn.is_displayed() and btn.is_enabled():
                    try:
                        btn.click()
                        break
                    except ElementClickInterceptedException:
                        driver.execute_script("arguments[0].click();", btn)
                        break
            time.sleep(1)

            for codigo, quantidade, descricao in itens_csv:
                try:
                    # Tenta buscar o item (com até 3 tentativas)
                    item_encontrado = self.buscar_item(driver, wait, codigo)
                    
                    if not item_encontrado:
                        motivo = "Item não encontrado após várias tentativas"
                        self.itens_pulados.append({'Codigo': codigo, 'Descricao': descricao, 'Quantidade': quantidade, 'Motivo': motivo})
                        continue


                    # Se o item foi encontrado, prossegue com a seleção
                    elemento_handlebar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[starts-with(@id, 'handlebar-') and .//span[normalize-space()='Selecionar']]")))
                    driver.execute_script("arguments[0].scrollIntoView(true);", elemento_handlebar)
                    time.sleep(0.5)
                    elemento_handlebar.click()
                    time.sleep(1)

                    qtd_input = wait.until(EC.presence_of_element_located((By.NAME, "QT_ESTOQUE")))
                    qtd_input.clear()
                    qtd_input.send_keys(quantidade)
                    time.sleep(0.5)

                    actions = ActionChains(driver)
                    for _ in range(5):
                        actions.send_keys(Keys.TAB)
                    actions.perform()
                    time.sleep(0.5)

                    actions = ActionChains(driver)
                    actions.send_keys(Keys.ENTER)
                    actions.perform()
                    time.sleep(3)

                except Exception as e_item:
                    motivo = str(e_item)
                    self.itens_pulados.append({'Codigo': codigo, 'Descricao': descricao, 'Quantidade': quantidade, 'Motivo': motivo})
                    continue

                    # APÓS PROCESSAR TODOS OS ITENS DO CSV, CLICA NOS BOTÕES FINAIS
            try:
                # Primeiro botão: Cancelar
                btn_cancelar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="handlebar-439428"]')))
                driver.execute_script("arguments[0].scrollIntoView(true);", btn_cancelar)
                time.sleep(0.5)
                btn_cancelar.click()
                time.sleep(1)
                
                # Segundo botão: via span
                elemento_span = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="handlebar-360728"]/span')))
                driver.execute_script("arguments[0].scrollIntoView(true);", elemento_span)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", elemento_span)
                time.sleep(2)  # Aguarda um pouco mais para a ação ser completada
                
            except TimeoutException:
                self.txt_pulados.insert(tk.END, "⚠️ Timeout: botões finais não encontrados\n")
            except Exception as e_final:
                self.txt_pulados.insert(tk.END, f"⚠️ Erro ao clicar nos botões finais: {e_final}\n")
                
            # MOSTRA MENSAGEM DE CONCLUÍDO SEM FECHAR O NAVEGADOR
            self.txt_pulados.insert(tk.END, "\n" + "="*150 + "\n")
            self.txt_pulados.insert(tk.END, "RESUMO FINAL:\n")
            self.txt_pulados.insert(tk.END, "="*150 + "\n\n")
            
            total_itens_csv = total_linhas_csv
            total_pulados = len(self.itens_pulados)
            total_repetidos = len(self.itens_repetidos)
            total_unicos = len(itens_csv)
            total_criados = total_unicos - len([i for i in self.itens_pulados if i['Codigo'] in itens_agrupados])
            total_ajustados = len(self.itens_ajustados)

            resumo = f"Total de linhas no CSV: {total_itens_csv}\n"
            resumo += f"Total de itens criados na requisição: {total_criados}\n"
            resumo += f"Itens repetidos ignorados: {total_repetidos}\n"
            resumo += f"Itens não encontrados: {total_pulados}\n"
            resumo += f"Itens com quantidade ajustada: {total_ajustados}\n\n"
            self.txt_pulados.insert(tk.END, resumo)

            # Mostra itens repetidos ignorados com fundo laranja
            if self.itens_repetidos:
                self.txt_pulados.insert(tk.END, "ITENS REPETIDOS IGNORADOS (mantido maior quantidade):\n", "laranja")
                self.txt_pulados.insert(tk.END, "-"*150 + "\n", "laranja")
                
                # Armazenar posições para aplicar a tag de fundo
                start_pos = self.txt_pulados.index("end-1c linestart")
                
                for item in self.itens_repetidos:
                    linha = f"Cod: {item['Codigo']} | "
                    linha += f" {item['Descricao'][:30]}... | "
                    if 'Quantidade_Nova' in item:
                        linha += f" {item['Quantidade_Antiga']} → {item['Quantidade_Nova']} | "
                    else:
                        linha += f" {item['Quantidade_Antiga']} (menor que {item['Quantidade_Atual_Maior']}) | "
                    linha += f" {item['Motivo']}\n"
                    self.txt_pulados.insert(tk.END, linha)
                
                # Aplicar fundo laranja a todo o bloco
                end_pos = self.txt_pulados.index("end-1c")
                self.txt_pulados.tag_add("laranja", start_pos, end_pos)
                self.txt_pulados.insert(tk.END, "\n")

            # Mostra itens ajustados com fundo amarelo
            if self.itens_ajustados:
                self.txt_pulados.insert(tk.END, "ITENS COM QUANTIDADE AJUSTADA:\n", "amarelo")
                self.txt_pulados.insert(tk.END, "-"*150 + "\n", "amarelo")
                
                # Armazenar posições para aplicar a tag de fundo
                start_pos = self.txt_pulados.index("end-1c linestart")
                
                for item in self.itens_ajustados:
                    linha = f"Cod: {item['Codigo']} | "
                    linha += f" {item['Descricao'][:30]}... | "
                    linha += f" {item['Quantidade_Original']} → "
                    linha += f" {item['Quantidade_Ajustada']} | "
                    linha += f" {item['Regra']}\n"
                    self.txt_pulados.insert(tk.END, linha)
                
                # Aplicar fundo amarelo a todo o bloco
                end_pos = self.txt_pulados.index("end-1c")
                self.txt_pulados.tag_add("amarelo", start_pos, end_pos)
                self.txt_pulados.insert(tk.END, "\n")

            # Mostra itens pulados com fundo vermelho
            if self.itens_pulados:
                self.txt_pulados.insert(tk.END, "ITENS PULADOS:\n", "vermelho")
                self.txt_pulados.insert(tk.END, "-"*150 + "\n", "vermelho")
                
                # Armazenar posições para aplicar a tag de fundo
                start_pos = self.txt_pulados.index("end-1c linestart")
                
                for item in self.itens_pulados:
                    linha = f"Código: {item['Codigo']} | "
                    linha += f"Descrição: {item['Descricao'][:30]}... | "
                    linha += f"Quantidade: {item['Quantidade']} | "
                    linha += f"Motivo: {item['Motivo']}\n"
                    self.txt_pulados.insert(tk.END, linha)
                
                # Aplicar fundo vermelho a todo o bloco
                end_pos = self.txt_pulados.index("end-1c")
                self.txt_pulados.tag_add("vermelho", start_pos, end_pos)
                self.txt_pulados.insert(tk.END, "\n")

            # Seção final com fundo verde
            start_pos = self.txt_pulados.index("end-1c")
            self.txt_pulados.insert(tk.END, "="*150 + "\n")
            self.txt_pulados.insert(tk.END, "PROCESSO CONCLUÍDO COM SUCESSO!\n")
            self.txt_pulados.insert(tk.END, f"Total de itens únicos processados: {total_unicos}\n")
            self.txt_pulados.insert(tk.END, f"Itens repetidos consolidados: {total_repetidos}\n")
            self.txt_pulados.insert(tk.END, "Todos os itens foram adicionados à requisição.\n")
            self.txt_pulados.insert(tk.END, "O navegador permanecerá aberto para revisão.\n")
            self.txt_pulados.insert(tk.END, "Feche manualmente quando desejar.\n")
            self.txt_pulados.insert(tk.END, "="*150 + "\n")
            end_pos = self.txt_pulados.index("end-1c")
            self.txt_pulados.tag_add("verde", start_pos, end_pos)
            
            self.txt_pulados.see(tk.END)

            # Mostra mensagem popup informando que terminou
            messagebox.showinfo("Processo Concluído", 
                               f"Requisição criada com sucesso!\n\n"
                               f"Total de itens no CSV: {total_itens_csv}\n"
                               f"Itens adicionados: {total_criados}\n"
                               f"Itens repetidos ignorados: {total_repetidos}\n"
                               f"Itens não encontrados: {total_pulados}\n"
                               f"Itens com quantidade ajustada: {total_ajustados}\n"
                               f"O navegador permanecerá aberto para revisão.")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro na automação:\n{e}")
            try:
                if driver:
                    driver.quit()
            except:
                pass

        self.btn_run.config(state=tk.NORMAL)

    def log(self, msg):
        pass

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()