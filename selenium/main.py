# std
from time import sleep
from os import chdir, remove, listdir
# externo
import pandas as Pd
import pyautogui as AutoGui
from selenium.webdriver.common.by import By
from selenium.webdriver import Edge, EdgeOptions
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager

# current working directory
CWD = "\\".join( __file__.split("\\")[0:-1] )

DRIVER = EdgeChromiumDriverManager().install()
OPTIONS = EdgeOptions()
OPTIONS.add_argument("--start-maximized")
# OPTIONS.add_argument("--headless")
OPTIONS.add_experimental_option("prefs", {
    "download.prompt_for_download": False,
    "download.default_directory": rf"{ CWD }\arquivos",
})

DE_PARA = {
    "labelFirstName": "First Name",
    "labelLastName": "Last Name",
    "labelCompanyName": "Company Name",
    "labelRole": "Role in Company",
    "labelAddress": "Address",
    "labelEmail": "Email",
    "labelPhone": "Phone Number",
}

def download_excel( navegador: Edge ) -> None:
    # clicar em download e esperar a página carregar
    abaOriginal = navegador.current_window_handle
    navegador.find_element(By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a")\
             .click()
    sleep(2)
    
    # clicar em "Baixar arquivo"
    AutoGui.click(831, 137)
    sleep(1)

    # retornar à aba do desafio
    navegador.switch_to.window(abaOriginal)

def main() -> None:
    navegador = Edge( OPTIONS, Service(DRIVER) )
    navegador.implicitly_wait(10)
    
    navegador.get("https://www.rpachallenge.com/")
    download_excel(navegador)

    # iniciar
    navegador.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button')\
             .click()

    # ler o Excel
    df = Pd.DataFrame( Pd.read_excel("./arquivos/challenge.xlsx") )
    df.columns = df.columns.str.strip() # remover espaços em branco das colunas
    linhasCsv: list[dict] = df.to_dict("records")

    # cada linha no csv
    for linha in linhasCsv:
        # cada input no form, exceto o Submit
        for _input in navegador.find_elements( By.XPATH, '//form/div//input' ):
            atributo = _input.get_attribute( "ng-reflect-name" )
            coluna = DE_PARA[atributo]
            _input.send_keys( linha[coluna] )
        
        # próxima página
        navegador.find_element( By.XPATH, '//form/input' )\
                 .click()
    
    # finalizar
    status = navegador.find_element(By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[2]/div[1]").text
    mensagem = navegador.find_element(By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[2]/div[2]").text
    AutoGui.alert(mensagem, status, "OK")
    navegador.quit()

"""
    Desafio: https://www.rpachallenge.com/
"""
if __name__ == "__main__":
    # altera o cwd para a pasta desse arquivo
    chdir(CWD)
    # Limpa a pasta dos arquivos
    [ remove(f"./arquivos/{ arquivo }") for arquivo in listdir("./arquivos") ]

    main()
    exit(0)