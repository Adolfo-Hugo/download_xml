import tkinter as tk
from tkinter import filedialog
import pandas as pd
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from tqdm import tqdm
import pyautogui


def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal do tkinter

    file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])
    if file_path:
        return file_path
    else:
        return None
 
def main():
    # Selecionar arquivo Excel
    excel = selecionar_arquivo()
    if not excel:
        print("Nenhum arquivo selecionado. Encerrando o programa.")
        return
    
    # Informar o caminho de salvamento do XML
    caminho = str(input('Informe o caminho de salvamento do XML: '))
    download_path = caminho
    
    # Configurar preferências do Chrome para download
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # Configurar o navegador Chrome
    
    navegador = webdriver.Chrome(service=Service(), options=chrome_options)
    link = "https://meudanfe.com.br"
    navegador.maximize_window()
    navegador.get(link)
    time.sleep(5)
    
    # Carregar dados do arquivo Excel
    try:
        chaves = pd.read_excel(excel)
    except Exception as e:
        print(f"Erro ao carregar o arquivo Excel: {e}")
        return
    
    # Iterar sobre as chaves do arquivo Excel
    pbar = tqdm(total=len(chaves), position=0, leave=True)
    for codigo_chave in chaves['chaves']:
        pbar.update()
        
        try:
            # Preencher a chave no campo correspondente
            navegador.find_element(By.XPATH,'//*[@id="get-danfe"]/div/div/div[1]/div/div/div/input').send_keys(codigo_chave)
            time.sleep(2)
            
            
            navegador.find_element(By.XPATH,'//*[@id="get-danfe"]/div/div/div[1]/div/div/div/button').click()
            time.sleep(2)
            
            # Esperar até que o botão de download esteja disponível
            while True:
                try:
                    botao = WebDriverWait(navegador, 20).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div/div[2]/button[1]'))
                    )
                    botao.click()
                    print(f'Chave N \033[32m{codigo_chave}\033[0m salva')
                    time.sleep(1)
                    # pyautogui.press('enter')
                    break
                except:
                    print("Esperando o botão de download estar disponível...")
                    time.sleep(5)  # Aguarde 5 segundos antes de tentar novamente
            
            time.sleep(3)

            # Renomear o arquivo baixado
            downloaded_file = max([f for f in os.listdir(download_path)], key=lambda x: os.path.getctime(os.path.join(download_path, x)))
            new_file_name = f"{codigo_chave}.xml"
            os.rename(os.path.join(download_path, downloaded_file), os.path.join(download_path, new_file_name))
            
            # Fechar o modal
            
            time.sleep(2)

            print('Procurando nova chave...')
            navegador.find_element(By.XPATH,'/html/body/div[1]/div/div[1]/div/div[1]/button').click()
            time.sleep(1)
        
        except Exception as e:
            print(f"Erro ao processar a chave {codigo_chave}: {e}")
    
    # Fechar o navegador ao finalizar
    navegador.quit()

if __name__ == "__main__":
    main()