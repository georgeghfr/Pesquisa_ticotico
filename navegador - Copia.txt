from selenium import webdriver
from selenium.webdriver.firefox.service import Service

# Pode manter a URL direto aqui se preferir
URL = "http://portal.servicos.ccarj.intraer/portal/faces/ticotico.xhtml"


def abrir_firefox():

    options = webdriver.FirefoxOptions()

    # ignora problemas de certificado
    options.accept_insecure_certs = True

    # Usa o geckodriver baixado manualmente na pasta do projeto
    service = Service(executable_path="./geckodriver.exe")

    driver = webdriver.Firefox(
        service=service,
        options=options
    )

    driver.maximize_window()

    driver.get(URL)

    return driver