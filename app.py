from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd

class Scrappy:

    def __init__(self):
        self.termo_pesquisa = ""

    def iniciar(self):
        self.termo_pesquisa = input("Digite o termo de pesquisa: ")
        self.pesquisa_google()

    def pesquisa_google(self):
        # Configurando o driver do Chrome
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service)
        self.driver.get('https://www.google.com.br/?h1=pt-BR')

        # Encontrar a barra de pesquisa do Google pelo nome do campo
        search_box = self.driver.find_element("name", "q")

        # Digitar o termo de pesquisa na barra de pesquisa
        search_box.send_keys(self.termo_pesquisa)
        
        sleep(2)
        
        search_box.send_keys(Keys.RETURN)

        # Esperar um pouco para os resultados serem carregados
        sleep(5)
        
        # Scroll down até o final da página
        self.scroll_to_bottom()

        # Coletar os dados e organizar em um DataFrame
        df = self.organizar_dados()
        
        # Exportar os dados para um arquivo Excel
        self.exportar_excel(df)
        
        # Fechar o navegador
        self.driver.quit()


    def scroll_to_bottom(self):
        # Obter a altura da página atual
        last_height = self.driver.execute_script("return document.body.scrollHeight")

        while True:
            # Scroll down para o final da página
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            # Aguardar o carregamento da página
            sleep(2)

            # Calcular nova altura e comparar com a altura anterior
            new_height = self.driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

    def organizar_dados(self):
        # Inicializar listas para armazenar os dados
        titles = []
        links = []
        snippets = []

        # Coletar os dados dos resultados da pesquisa
        search_results = self.driver.find_elements(By.XPATH, '//div[@jscontroller="SC7lYd"]')
        for result in search_results:
            # Capturar título, URL e snippet
            title = result.find_element(By.XPATH, './/h3').text
            link = result.find_element(By.XPATH, './/a').get_attribute('href')
            snippet = result.find_element(By.XPATH, './/span').text

            # Armazenar os dados nas listas
            titles.append(title)
            links.append(link)
            snippets.append(snippet)

        # Criar um DataFrame do Pandas com os dados coletados
        df = pd.DataFrame({
            'Title': titles,
            'Link': links,
            'Snippet': snippets
        })

        return df
    
    def exportar_excel(self, df):
        # Especificar o nome do arquivo Excel
        excel_file = f'pesquisa_{self.termo_pesquisa}.xlsx'
        
        # Exportar o DataFrame para Excel
        df.to_excel(excel_file, index=False)
        print(f'Dados exportados para {excel_file}')

start = Scrappy()
start.iniciar()