import os
import shutil
import json
import logging
import csv
import re
import time
import datetime
import xml.etree.ElementTree as ET
import win32api
import tkinter as tk 
from tkinter import filedialog, messagebox, ttk
import threading
import time

ultima_impressao = []
impressao_realizada = False
log_directory = 'C:\\Log de Laudos' # Configuração do sistema de log
log_file_path = os.path.join(log_directory, 'app.logLaudos')

if not os.path.exists(log_directory):# Verifica se o diretório de log existe, se não, cria o diretório
    os.makedirs(log_directory)

logging.basicConfig(
    filename=log_file_path,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%d/%m/%Y %H:%M:%S'  # Formato de data e hora
)
CONFIG_FILE = "config.json"
DEFAULT_CONFIG = {'diretorio_origem_xml': 'C:\\Caminho\\Padrao\\Para\\XML',
                  'diretorio_destino_xml': 'C:\\Caminho\\Padrao\\Para\\Destino',
                  'diretorio_laudos': 'C:\\Caminho\\Padrao\\Para\\Laudos'}
CSV_HEADER = ['Data', 'Chave', 'Produto', 'Lote', 'Tipo', 'Laudo_Impresso']
CSV_FILE = os.path.join(log_directory, 'log.csv')
diretorio_origem_xml = ""# Variáveis para armazenar os diretórios selecionados pelo usuário
diretorio_destino_xml = ""# Variáveis para armazenar os diretórios selecionados pelo usuário
diretorio_laudos = ""# Variáveis para armazenar os diretórios selecionados pelo usuário

def iniciar_csv():
    if not os.path.isfile(CSV_FILE):# Cria o arquivo CSV com o cabeçalho se o arquivo não existir
        with open(CSV_FILE, 'w', newline='', encoding='utf-8-sig') as file:
            writer = csv.writer(file, delimiter=';')
            writer.writerow(CSV_HEADER)

def registrar_log(chave, produto='', lote='', tipo='', laudo_impresso=''):
    data_atual = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    with open(CSV_FILE, 'a', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow([data_atual, chave, produto, lote, tipo, laudo_impresso])

def abrir_csv():# Tenta abrir o arquivo CSV com o programa padrão do sistema    
    if os.path.isfile(CSV_FILE):
        os.startfile(CSV_FILE)  # Funciona apenas no Windows
    else:
        tk.messagebox.showerror("Erro", "O arquivo de log não existe. Crie antes de tentar abri-lo.")

def verificar_existencia_xml(caminho_xml):
    if os.path.exists(caminho_xml):
        return True
    else:
        logging.warning(f"O XML {caminho_xml} não foi encontrado.")
        return False

def extrair_dados_lotes(caminho_xml, printer_name):
    tree = ET.parse(caminho_xml)
    root = tree.getroot()
    namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
    chave = os.path.splitext(os.path.basename(caminho_xml))[0]
    lotes_produtos = []  # Lista para armazenar todos os lotes e produtos encontrados
    dados_impressao = {}  # Dicionário para rastrear o status de cada impressão
    for det in root.findall(".//nfe:det", namespaces):
        produto = det.find("./nfe:prod/nfe:xProd", namespaces).text
        lote = det.find("./nfe:prod/nfe:rastro/nfe:nLote", namespaces).text
        ncm = det.find("./nfe:prod/nfe:NCM", namespaces).text  # Busca o NCM
        laudo_impresso = False  # Flag para rastrear se o laudo foi impresso com sucesso        
        if ncm.startswith(("2925", "3001", "3002", "3003", "3004", "3005", "3306", "3006", "3007", "3824", "1512",
                           "2106", "2207", "2511", "2712", "2828", "2835", "2847", "2909", "2912", "2915", "2922",
                           "2932", "2933", "2935", "2939", "2941", "3204", "3304", "3307", "3401", "3402", "3504",
                           "3507", "3808", "7017", "8419", "9018")):  # Verifica o NCM antes de processar (considerando 3002, 3004, 3005 e 3006 e demais NCMS)
            lotes_produtos.append((produto, lote))
            laudo = buscar_laudo(produto, lote)
            if laudo:
                try:
                    win32api.ShellExecute(0, "print", laudo, f'/d:"{printer_name}"', ".", 0)
                    laudo_impresso = True
                    registrar_log(chave=chave, produto=produto, lote=lote, tipo="INFO",
                                   laudo_impresso="Impresso com sucesso!")
                    time.sleep(1)  # Adiciona um atraso de 1 segundo entre as impressões
                except Exception as e:
                    registrar_log(chave=chave, produto=produto, lote=lote, tipo="ERROR",
                                   laudo_impresso=f"Erro ao imprimir: {str(e)}")
            else:
                registrar_log(chave=chave, produto=produto, lote=lote, tipo="ERROR",
                               laudo_impresso="Laudo não encontrado.")
            dados_impressao[(produto, lote)] = {
                'ncm': ncm,
                'laudo_impresso': laudo_impresso,
                'mensagem_erro': None if laudo_impresso else "Laudo não encontrado ou erro ao imprimir"
            }  # Adiciona o status da impressão ao dicionário    
    if not lotes_produtos:  # Log geral se nenhum lote ou produto foi encontrado
        registrar_log(chave=chave, tipo="INFO", laudo_impresso="Nenhum lote ou produto encontrado no XML.")
    else:       
        mensagem = "Detalhes dos lotes e produtos:\n"  # Compila a mensagem de detalhes dos lotes e produtos
        for produto, lote in lotes_produtos:
            mensagem += f"Produto: {produto}, Lote: {lote}\n"            
            status_impresso = dados_impressao[(produto, lote)]['laudo_impresso']  # Adiciona detalhes sobre a impressão de cada laudo
            mensagem += f"Impresso: {'Sim' if status_impresso else 'Não'}\n"
        registrar_log(chave=chave, tipo="INFO", laudo_impresso=mensagem)
    return dados_impressao  # Retorna um dicionário com informações sobre a impressão de cada laudo

def buscar_laudo(produto, Lote):
    diretorio_laudos = obter_diretorio_laudos()
    lote_corrigido = substituir_caracteres(Lote)
    for root, dirs, files in os.walk(diretorio_laudos):
        for file in files:
            if lote_corrigido.lower() in file.lower() and file.endswith(".pdf"):
                return os.path.join(root, file)
    return None

def substituir_caracteres(nome_arquivo):
    return re.sub(r'/', '-', nome_arquivo)

def abrir_selecao_diretorio_origem():
    global diretorio_origem_xml
    diretorio_origem_xml = filedialog.askdirectory()
    if diretorio_origem_xml:
        entry_origem.delete(0, tk.END)
        entry_origem.insert(0, diretorio_origem_xml)
        messagebox.showinfo("Sucesso", f"Diretório de origem definido para: {diretorio_origem_xml}")
        salvar_diretorios()

def abrir_selecao_diretorio_destino():
    global diretorio_destino_xml
    diretorio_destino_xml = filedialog.askdirectory()
    if diretorio_destino_xml:
        entry_destino.delete(0, tk.END)
        entry_destino.insert(0, diretorio_destino_xml)
        messagebox.showinfo("Sucesso", f"Diretório de destino definido para: {diretorio_destino_xml}")
        salvar_diretorios()

def abrir_selecao_diretorio_laudos():
    global diretorio_laudos
    diretorio_laudos = filedialog.askdirectory()
    if diretorio_laudos:
        entry_laudos.delete(0, tk.END)
        entry_laudos.insert(0, diretorio_laudos)
        messagebox.showinfo("Sucesso", f"Diretório de laudos definido para: {diretorio_laudos}")
        salvar_diretorios()

def salvar_diretorios():
    config = {
        'diretorio_origem_xml': diretorio_origem_xml,
        'diretorio_destino_xml': diretorio_destino_xml,
        'diretorio_laudos': diretorio_laudos
    }
    with open(CONFIG_FILE, 'w') as config_file:
        json.dump(config, config_file)
        messagebox.showinfo("Diretórios", f"Sucesso!, Diretórios Salvos")

def renomear_mover_xmls():
    if not diretorio_origem_xml or not diretorio_destino_xml:
        messagebox.showerror("Erro",
                             "Por favor, selecione ambos os diretórios de origem e destino antes de prosseguir.")
        return
    arquivos_transferidos = []
    arquivos_com_erro = []
    arquivos_validos = False  # Flag para indicar se existem arquivos válidos para transferência
    for filename in os.listdir(diretorio_origem_xml):
        if filename.endswith("-nfe.xml"):
            arquivos_validos = True  # Existem arquivos válidos para transferência
            xml_file = os.path.join(diretorio_origem_xml, filename)
            if os.path.isfile(xml_file):
                try:
                    tree = ET.parse(xml_file)
                    root = tree.getroot()
                    nNF_element = root.find(".//{http://www.portalfiscal.inf.br/nfe}nNF")
                    if nNF_element is not None and nNF_element.text is not None:
                        nNF_value = nNF_element.text
                        new_file_name = f"{nNF_value}.xml"
                        new_file_path = os.path.join(diretorio_destino_xml, new_file_name)
                        shutil.copy2(xml_file, new_file_path)
                        os.remove(xml_file)
                        registrar_log(f"Arquivo movido para: {new_file_path}")
                        arquivos_transferidos.append(filename)
                    else:
                        logging.warning(f"Elemento nNF não encontrado ou sem texto em: {xml_file}")
                        arquivos_com_erro.append(filename)
                except ET.ParseError as parse_error:
                    logging.error(f"Erro ao fazer parse do XML {xml_file}: {parse_error}")
                    arquivos_com_erro.append(filename)
                except Exception as e:
                    logging.error(f"Erro desconhecido ao processar o arquivo {xml_file}: {e}")
                    arquivos_com_erro.append(filename)
            else:
                logging.warning(f"{xml_file} não é um arquivo")
                arquivos_com_erro.append(filename)
    if not arquivos_validos:
        messagebox.showinfo("Aviso", "Não há arquivos a serem transferidos")
        return
    mensagem = "Arquivos transferidos com sucesso:\n\n" if arquivos_transferidos else ""
    mensagem += "\n".join(arquivos_transferidos)
    if arquivos_com_erro:
        mensagem += "\n\nArquivos com erro:\n\n"
        mensagem += "\n".join(arquivos_com_erro)
        messagebox.showerror("Erro", f"Erro ao transferir alguns arquivos:\n\n{mensagem}")
    else:
        messagebox.showinfo("Sucesso", f"Arquivos transferidos com sucesso:\n\n{mensagem}")
    time.sleep(3)

def imprimir():
    global impressao_cancelada, impressao_realizada
    if not impressao_cancelada:
        if not impressao_realizada:
            numeros_xml = solicitar_numeros_xml()
            printer_name = "Nome_da_Impressora"  # Defina o nome da impressora aqui
            notas_fiscais = imprimir_lotes_xml(numeros_xml, printer_name)
            if not notas_fiscais:
                messagebox.showinfo("Aviso", "Nenhuma nota fiscal encontrada.")
            impressao_realizada = True
        else:
            messagebox.showinfo("Aviso", "A impressão já foi realizada.")
        impressao_cancelada = True
    else:
        messagebox.showinfo("Aviso", "A impressão já foi cancelada.")

def exibir_mensagem_confirmacao(total_laudos_impressos):
    mensagem = f"Todos os laudos foram impressos. Total de laudos impressos: {total_laudos_impressos}\nDeseja recomeçar?"
    resposta = messagebox.askyesno("Confirmação", mensagem)
    if resposta:
        renomear_mover_xmls()

def imprimir():
    global impressao_cancelada
    if not impressao_cancelada:
        numeros_xml = solicitar_numeros_xml()
        notas_fiscais = imprimir_lotes_xml(numeros_xml)
        if not notas_fiscais:
            messagebox.showinfo("Aviso", "Nenhuma nota fiscal encontrada.")
        impressao_cancelada = True
    else:
        messagebox.showinfo("Aviso", "A impressão já foi cancelada.")

def imprimir_lotes_xml(numeros_xml):
    global impressao_realizada
    impressao_realizada = False  # Reiniciar a variável de controle
    printer_name = "192.168.1.33"  # Defina seu nome de impressora aqui
    laudos_impressos = []
    laudos_faltantes = []
    total_impresso = 0
    for numero_xml in numeros_xml:
        caminho_xml = fr"{diretorio_destino_xml}\\{numero_xml}.xml"
        if verificar_existencia_xml(caminho_xml):
            dados_notas, laudos = extrair_dados_lotes(caminho_xml, printer_name)
            if dados_notas:  # Assumindo que dados_notas contém os resultados das impressões
                for chave, info in dados_notas.items():
                    if info['impresso']:  # Supondo que cada laudo tem um status 'impresso'
                        laudos_impressos.append(info['produto'])
                        total_impresso += 1
                    else:
                        laudos_faltantes.append(info['produto'])
    impressao_realizada = True
    exibir_mensagem_confirmacao(total_impresso, laudos_faltantes)

def exibir_mensagem_confirmacao(total_impresso, laudos_faltantes):
    if total_impresso > 0 and not laudos_faltantes:
        mensagem = f"Todos os {total_impresso} laudos foram impressos com sucesso."
        messagebox.showinfo("Sucesso", mensagem)
    elif laudos_faltantes:
        mensagem = f"Alguns laudos não foram impressos. Verifique os seguintes produtos: {', '.join(laudos_faltantes)}"
        messagebox.showerror("Erro", mensagem)
    else:
        messagebox.showinfo("Informação", "Nenhum laudo foi impresso.")

def exibir_mensagem_confirmacao(total_laudos_impressos):
    mensagem = f"Todos os laudos foram impressos. Total de laudos impressos: {total_laudos_impressos}\nDeseja recomeçar?"
    resposta = messagebox.askyesno("Confirmação", mensagem)
    if resposta:
        renomear_mover_xmls()

def reiniciar_processo():
    if not diretorio_destino_xml:
        messagebox.showerror("Erro", "Por favor, selecione o diretório de destino antes de prosseguir.")
        return
    renomear_mover_xmls()

def obter_diretorio_laudos():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as config_file:
            try:
                config = json.load(config_file)
                return config.get('diretorio_laudos', DEFAULT_CONFIG['diretorio_laudos'])
            except json.JSONDecodeError:
                logging.error(
                    f"Erro ao carregar o arquivo JSON. Usando diretório padrão: {DEFAULT_CONFIG['diretorio_laudos']}")
                return DEFAULT_CONFIG['diretorio_laudos']
    else:
        logging.warning(
            f"Arquivo de configuração não encontrado. Usando diretório padrão: {DEFAULT_CONFIG['diretorio_laudos']}")
        return DEFAULT_CONFIG['diretorio_laudos']

def salvar_diretorio_laudos(diretorio_laudos):
    with open(CONFIG_FILE, 'w') as config_file:
        json.dump({'diretorio_laudos': diretorio_laudos}, config_file)

def mostrar_ajuda():
    mensagem = """
    Para imprimir lotes da Nota Fiscal, insira os números das Notas fiscais desejados, separados por vírgula ou intervalo.
    Por exemplo:
     - 1,2,3 (use para notas fiscais específicas)
     - 5-10 (use para notas fiscais em sequência)
     - O botão de transferência de xml serve para que a cada nota faturada no sistema o arquivo xml seja transferido para o nosso drive e seja possivel a impressão dos laudos
    """
    messagebox.showinfo("Ajuda", mensagem)

def get_log_path():  # Substitua com o caminho onde você quer salvar o arquivo de log
    return "C:\\Log de Laudos\\log.csv"

def gerar_csv_log():
    log_path = get_log_path()    
    if not os.path.exists(log_path):# Cria o cabeçalho se o arquivo não existir
        with open(log_path, 'w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.writer(csv_file, delimiter=';')  # Usando delimitador ';'
            writer.writerow(CSV_HEADER)
    try:
        with open(log_file_path, 'r') as log_file, open(log_path, 'a', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.writer(csv_file, delimiter=';')
            for line in log_file:                
                data = line.strip().split(' - ', 2)  # Dividindo por nível de log e mensagem # Assumindo que o log é registrado no formato especificado pelo logging
                if len(data) == 3:
                    timestamp = data[0]
                    nivel = data[1]
                    mensagem = data[2]
                    writer.writerow([timestamp, nivel, "Log", mensagem])
        messagebox.showinfo("Sucesso", "Arquivo CSV de log gerado com sucesso! em " + log_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar arquivo CSV de log: {e}")

def solicitar_numeros_xml():
    numeros = entry_xml.get()
    numeros = numeros.replace(" ", "")
    numeros = numeros.split(",")
    numeros_xml = []
    for numero in numeros:
        if "-" in numero:
            inicio, fim = map(int, numero.split("-"))
            numeros_xml.extend(range(inicio, fim + 1))
        else:
            numeros_xml.append(int(numero))
    return numeros_xml
if os.path.exists(CONFIG_FILE):# Carregar diretórios salvos
    with open(CONFIG_FILE, 'r') as config_file:
        config = json.load(config_file)
        diretorio_origem_xml = config.get('diretorio_origem_xml', DEFAULT_CONFIG['diretorio_origem_xml'])
        diretorio_destino_xml = config.get('diretorio_destino_xml', DEFAULT_CONFIG['diretorio_destino_xml'])
        diretorio_laudos = config.get('diretorio_laudos', DEFAULT_CONFIG['diretorio_laudos'])

def mostrar_ultimos_laudos_impressos():
    try:
        # Caminho do arquivo CSV de log        
        log_file_path = "C:\\Log de Laudos\\log.csv"  # Atualize o caminho conforme sua estrutura        
        if not os.path.isfile(log_file_path):  # Verificar se o arquivo CSV existe
            messagebox.showwarning("Aviso", "O arquivo de log não existe.")
            return       

        # Abrir o arquivo CSV de log para leitura
        with open(log_file_path, 'r', newline='', encoding='utf-8-sig') as file:
            reader = csv.DictReader(file, delimiter=';')

            # Inicializar a lista para armazenar a última linha válida
            ultimos_laudos_nf = []  

            # Iterar pelas linhas de trás para frente
            for row in reversed(list(reader)):  # Leitura invertida das linhas do CSV
                if row['Laudo_Impresso']:  # Verificar se o laudo foi impresso
                    chave_nf = row['Chave']
                    produto = row['Produto']
                    lote = row['Lote']
                    tipo = row['Tipo']
                    laudo_impresso = row['Laudo_Impresso']
                    linha_formatada = (row['Data'], chave_nf, produto, lote, tipo, laudo_impresso)
                    ultimos_laudos_nf.append(linha_formatada)
                    break  # Para após encontrar o primeiro laudo impresso

        # Exibir os laudos da última NF em forma de tabela
        if ultimos_laudos_nf:
            root_table = tk.Tk()
            root_table.title("Últimos Laudos Impressos")               

            # Criar uma tabela usando ttk.Treeview
            table = ttk.Treeview(root_table, columns=('Data', 'Chave', 'Produto', 'Lote', 'Tipo', 'Laudo Impresso'),
                                 show='headings')                

            style = ttk.Style()  # Definir estilo para a tabela
            style.theme_use("default")
            style.configure("Treeview", background="#D3D3D3", foreground="black", rowheight=25,
                            fieldbackground="#D3D3D3")
            style.map("Treeview", background=[('selected', '#347083')])                
            style.configure("Treeview.Heading", font=('Helvetica', 10, 'bold'), background="#4CAF50",
                            foreground="white")                

            scrollbar_y = ttk.Scrollbar(root_table, orient="vertical", command=table.yview)  # Configurar scrollbar
            scrollbar_y.pack(side="right", fill="y")
            table.configure(yscrollcommand=scrollbar_y.set)

            # Cabeçalhos para a tabela
            table.heading('Data', text='Data')
            table.heading('Chave', text='Nota')
            table.heading('Laudo Impresso', text='Laudo Impresso')
            table.column('Laudo Impresso', stretch=True, width=1000)

            # Adicionar os laudos à tabela
            for laudo in ultimos_laudos_nf:
                table.insert('', 'end', values=laudo)

            table.pack(expand=True, fill='both')
            root_table.mainloop()
        else:
            messagebox.showinfo("Últimos Laudos Impressos", "Nenhum laudo impresso na última NF.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo de log: {e}")

def verificar_resultados_impressao(dados_impressao):
    total_impressos = sum(1 for info in dados_impressao.values() if info['laudo_impresso'])
    total_erros = len(dados_impressao) - total_impressos

    if total_erros == 0 and total_impressos > 0:
        messagebox.showinfo("Resultado da Impressão", f"Todos os {total_impressos} laudos foram impressos com sucesso.")
    elif total_erros > 0:
        mensagem = f"{total_impressos} laudos foram impressos com sucesso, mas {total_erros} não puderam ser impressos. Verifique em ultimos laudos impressos para mais detalhes."
        messagebox.showwarning("Resultado da Impressão", mensagem)
    else:
        messagebox.showinfo("Resultado da Impressão", "Nenhum laudo foi impresso.")

def imprimir_lotes_xml(numeros_xml, printer_name):# Suponha que esta seja a função onde você processa cada XML    
    for numero_xml in numeros_xml:
        caminho_xml = fr"{diretorio_destino_xml}\\{numero_xml}.xml"
        if verificar_existencia_xml(caminho_xml):
            dados_impressao = extrair_dados_lotes(caminho_xml, printer_name)
            verificar_resultados_impressao(dados_impressao)

def imprimir():
    numeros_xml = solicitar_numeros_xml()
    if not numeros_xml:
        messagebox.showinfo("Aviso", "Nenhum número de XML selecionado.")
        return
    printer_name = "Nome_da_Impressora"  # Defina o nome da impressora aqui
    notas_fiscais = imprimir_lotes_xml(numeros_xml, printer_name)

def solicitar_numeros_xml():
    numeros = entry_xml.get()
    numeros = numeros.replace(" ", "")
    numeros = numeros.split(",")
    numeros_xml = []
    for numero in numeros:
        if "-" in numero:
            inicio, fim = map(int, numero.split("-"))
            numeros_xml.extend(range(inicio, fim + 1))
        else:
            numeros_xml.append(int(numero))
    return numeros_xml

root = tk.Tk()
root.title("Sistema de Impressão de Laudos")

frame_diretorios = tk.LabelFrame(root, text="Diretórios")
frame_diretorios.pack(padx=10, pady=10, fill="both", expand="yes")

label_origem = tk.Label(frame_diretorios, text="Diretório Origem XML:")
label_origem.grid(row=0, column=0, sticky="w")

entry_origem = tk.Entry(frame_diretorios, width=50)
entry_origem.grid(row=0, column=1, padx=5, pady=5, sticky="we")
entry_origem.insert(0, diretorio_origem_xml)

btn_origem = tk.Button(frame_diretorios, text="Selecionar", command=abrir_selecao_diretorio_origem)
btn_origem.grid(row=0, column=2, padx=5, pady=5)

label_destino = tk.Label(frame_diretorios, text="Diretório Destino XML:")
label_destino.grid(row=1, column=0, sticky="w")

entry_destino = tk.Entry(frame_diretorios, width=50)
entry_destino.grid(row=1, column=1, padx=5, pady=5, sticky="we")
entry_destino.insert(0, diretorio_destino_xml)

btn_destino = tk.Button(frame_diretorios, text="Selecionar", command=abrir_selecao_diretorio_destino)
btn_destino.grid(row=1, column=2, padx=5, pady=5)

label_laudos = tk.Label(frame_diretorios, text="Diretório Laudos:")
label_laudos.grid(row=2, column=0, sticky="w")

entry_laudos = tk.Entry(frame_diretorios, width=50)
entry_laudos.grid(row=2, column=1, padx=5, pady=5, sticky="we")
entry_laudos.insert(0, diretorio_laudos)

btn_laudos = tk.Button(frame_diretorios, text="Selecionar", command=abrir_selecao_diretorio_laudos)
btn_laudos.grid(row=2, column=2, padx=5, pady=5)

frame_xml = tk.LabelFrame(root, text="Impressão de Laudos")
frame_xml.pack(padx=10, pady=10, fill="both", expand="yes")

label_xml = tk.Label(frame_xml, text="Números das Notas Fiscais (Separados por vírgula ou intervalo):")
label_xml.grid(row=0, column=0, padx=5, pady=5, sticky="w")

entry_xml = tk.Entry(frame_xml, width=50)
entry_xml.grid(row=0, column=1, padx=5, pady=5, sticky="we")

btn_imprimir = tk.Button(frame_xml, text="Imprimir Laudos", command=imprimir)
btn_imprimir.grid(row=0, column=2, padx=5, pady=5)

frame_opcoes = tk.LabelFrame(root, text="Opções")
frame_opcoes.pack(padx=10, pady=10, fill="both", expand="yes")

btn_reiniciar = tk.Button(frame_opcoes, text="Transferir Xml", command=reiniciar_processo)
btn_reiniciar.grid(row=0, column=0, padx=5, pady=5)

btn_gerar_csv_log = tk.Button(frame_opcoes, text="Gerar CSV de Log", command=gerar_csv_log)
btn_gerar_csv_log.grid(row=0, column=1, padx=5, pady=5)

btn_ajuda = tk.Button(frame_opcoes, text="Ajuda", command=mostrar_ajuda)
btn_ajuda.grid(row=0, column=5, padx=5, pady=5)

button_log = tk.Button(frame_opcoes, text="Abrir log", command=abrir_csv)
button_log.grid(row=0, column=2, padx=10, pady=5)

btn_ultimos_laudos = tk.Button(frame_opcoes, text="Últimos Laudos Impressos", command=mostrar_ultimos_laudos_impressos)
btn_ultimos_laudos.grid(row=0, column=3, padx=5, pady=5)

root.mainloop()