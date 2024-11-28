from tkinter import Tk, Label, Entry, Button, StringVar, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import openpyxl
from time import sleep

def extrair_dados(numero_oab, uf):
    # Carregar planilha
    planilha_dados_consulta = openpyxl.load_workbook('dados_de_processos.xlsx')
    pagina_processos = planilha_dados_consulta['processos']
    
    # 1. Entrar no site
    driver = webdriver.Chrome()
    driver.get('https://pje-consulta-publica.tjmg.jus.br/')
    sleep(5)
    
    try:
        # 2. Inserir o número da OAB
        campo_numero_oab = driver.find_element(By.XPATH, "//input[@id='fPP:Decoration:numeroOAB']")
        campo_numero_oab.send_keys(numero_oab)
        
        # 3. Selecionar UF
        selecao_uf = driver.find_element(By.XPATH, "//select[@id='fPP:Decoration:estadoComboOAB']")
        opcoes_uf = Select(selecao_uf)
        opcoes_uf.select_by_visible_text(uf)
        
        # 4. Clicar em Pesquisar
        botao_pesquisar = driver.find_element(By.XPATH, "//input[@id='fPP:searchProcessos']")
        botao_pesquisar.click()
        sleep(5)
        
        # 5. Extrair dados dos processos
        links_abrir_processo = driver.find_elements(By.XPATH, "//a[@title='Ver Detalhes']")
        
        for link in links_abrir_processo:
            janela_principal = driver.current_window_handle
            link.click()
            sleep(5)
            janelas_abertas = driver.window_handles
            for janela in janelas_abertas:
                if janela != janela_principal:
                    driver.switch_to.window(janela)
                    sleep(5)
                    numero_processo = driver.find_elements(By.XPATH, "//div[@class='propertyView ']//div[@class='col-sm-12 ']")[0]
                    participantes = driver.find_elements(By.XPATH, "//tbody[contains(@id,'processoPartesPoloAtivoResumidoList:tb')]//span[@class='text-bold']")
                    lista_participantes = [participante.text for participante in participantes]
                    
                    # Armazenar dados na planilha
                    if lista_participantes:
                        pagina_processos.append([numero_oab, numero_processo.text, ', '.join(lista_participantes)])
                    
                    planilha_dados_consulta.save('dados_de_processos.xlsx')
                    driver.close()
                    
            driver.switch_to.window(janela_principal)
    finally:
        driver.quit()
        messagebox.showinfo("Sucesso", "Dados extraídos e salvos com sucesso!")

# Função para iniciar a extração com base nos inputs da interface
def iniciar_extracao():
    numero_oab = entrada_oab.get()
    uf = entrada_uf.get()
    if numero_oab and uf:
        extrair_dados(numero_oab, uf)
    else:
        messagebox.showwarning("Atenção", "Preencha todos os campos!")

# Interface Tkinter
app = Tk()
app.title("Consulta de Processos")

# Rótulos e Entradas
Label(app, text="Número da OAB:").pack()
entrada_oab = Entry(app)
entrada_oab.pack()

Label(app, text="UF:").pack()
entrada_uf = Entry(app)
entrada_uf.pack()

# Botão para iniciar extração
Button(app, text="Consultar", command=iniciar_extracao).pack()

# Iniciar a aplicação
app.mainloop()
