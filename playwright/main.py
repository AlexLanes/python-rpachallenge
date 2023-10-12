# std
from time import sleep
# externo
import pandas as Pd
import pyautogui as AutoGui
from playwright.sync_api import sync_playwright, Browser, Page

def obter_campo_excel( nomeInput: str ) -> str:
    match nomeInput:
        case "labelRole": return "Role in Company"
        case "labelPhone": return "Phone Number"
        case "labelEmail": return "Email"
        case "labelAddress": return "Address"
        case "labelLastName": return "Last Name"
        case "labelFirstName": return "First Name"
        case "labelCompanyName": return "Company Name"
        case _: raise Exception("Nome input inesperado")

def main( navegador: Browser, pagina: Page ) -> None:
    pagina.goto("https://www.rpachallenge.com/")
    
    # download excel
    # salvar na pasta de downloads
    nomeArquivo: str
    with pagina.expect_download() as download:
        cordenadasDownload = AutoGui.locateCenterOnScreen( "./screenshots/download.png", confidence=0.8 )
        AutoGui.click(cordenadasDownload)
        
        arquivo = download.value
        nomeArquivo = arquivo.suggested_filename
        arquivo.save_as(f"./downloads/{nomeArquivo}")
    
    # leitura e parse do arquivo excel
    df = Pd.DataFrame( Pd.read_excel(f"./downloads/{nomeArquivo}") )
    df.columns = df.columns.str.strip() # remover espaços em branco das colunas
    excel: list[dict] = df.to_dict("records")

    # clicar em iniciar
    cordenadasStart = AutoGui.locateCenterOnScreen( "./screenshots/start.png", confidence=0.8 )
    AutoGui.click(cordenadasStart)
    
    # percorrer todas as linhas do excel
        # percorrer todos os inputs do form atual que serão preenchidos
            # obter o ng-reflect-name do inpunt
            # obter o valor da coluna do excel com base no ng-reflect-name
            # preencher o input
        # clicar em submit
    for linha in excel:
        inputs = pagina.locator("xpath=/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/div//input")
        for index in range(0, inputs.count(), 1):
            nomeInput = inputs.nth(index).get_attribute("ng-reflect-name")
            nomeCampoExcel = obter_campo_excel(nomeInput)
            valorCampoExcel = linha[nomeCampoExcel]
            inputs.nth(index).fill( str(valorCampoExcel) )
        pagina.locator("xpath=/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input").click()

    # mostrar mensagem
    mensagem = pagina.locator("xpath=/html/body/app-root/div[2]/app-rpa1/div/div[2]/div[2]").inner_text()
    AutoGui.alert(mensagem, "Mensagem", "Finalizar")


if __name__ == "__main__":
    with sync_playwright() as playwright: 
        navegador = playwright.chromium.launch( channel="msedge", headless=False, args=["--start-maximized"] )
        navegador.new_context()
        pagina = navegador.new_page( no_viewport=True )
        main( navegador, pagina )
        navegador.close()
        exit(0)