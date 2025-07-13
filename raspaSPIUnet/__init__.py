#!/usr/bin/env python3
# coding: utf-8
'''
@author: Lauro Sá <memoriadcalculo@gmail.com>

Raspador para o site do Sistema de Gerenciamento dos Imóveis de Uso Especial (SPIUnet),
do Governo Federal.

A seguinte fonte foi utilizada para código referente ao LibreOffice:
https://christopher5106.github.io/office/2015/12/06/openoffice-libreoffice-automate-your-office-tasks-with-python-macros.html
'''
from __future__ import unicode_literals

# Configura o local para padrão de números, datas, etc
import locale
locale.setlocale(locale.LC_ALL, '')

# Verifica se está rodando no LibreOffice
try:
    import uno
    from com.sun.star.script.provider import XScript
    from scriptforge                  import CreateScriptService
    LO = True
except ImportError:
    LO = False

# Importa os módulos básicos necessários
# import csv
# import os
# import platform
# import ssl
# import subprocess
# import sys
# from datetime                     import datetime, date
# from pathlib                      import Path
# from threading                    import Thread
# from time                         import sleep
# from urllib.parse                 import urlunparse, urlparse
# from urllib.request               import build_opener, install_opener, ProxyHandler, urlretrieve

# TODO: pacotes externos
from selenium.webdriver                import Chrome
from selenium.webdriver                import ChromeOptions
from selenium.webdriver.chrome.service import Service
from bs4                                            import BeautifulSoup
from selenium.common.exceptions                     import TimeoutException
from selenium.webdriver.support.ui                  import WebDriverWait
from selenium.webdriver.support.expected_conditions import visibility_of_element_located
from selenium.webdriver.common.by      import By

def BASE_DIR():
    """
    Define o caminho inicial do script/planilha.
    """
    from pathlib                      import Path
    
    global LO
    
    if LO:
        doc         = XSCRIPTCONTEXT.getDocument()
        caminho_str = doc.URL
        caminho_sys = uno.fileUrlToSystemPath(caminho_str)
        caminho_obj = Path(caminho_sys)
        resultado   = caminho_obj.parent
    else:
        resultado   = Path(__file__).resolve().parent.parent
    
    return resultado

def BASE_PY():
    """
    Procura o caminho do executável do python.
    Necessário porque 'sys.executable' retorna o padrão do sistema e não do script.
    """
    import os
    import sys
    
    resultado = ""
    for sysPasta in sys.path:
        if sys.platform == "win32":
            resultado = os.path.join(sysPasta, "python.exe")
        else:
            resultado = os.path.join(sysPasta, "python")
        if os.path.isfile(resultado):
            break
    if not os.path.isfile(resultado):
        resultado = sys.executable
    
    return resultado

def LO_ScriptBasic(macro='Atualiza', module='raspaSPIUnet', library='Standard', noDoc=True):
    """
    Executa script Basic do LibreOffice.
    """
    
    ctx  = uno.getComponentContext()
    smgr = ctx.ServiceManager
    if noDoc:
        ins_ctx   = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
        scriptPro = ins_ctx.CurrentComponent.getScriptProvider()
        location  = "document"
    else:
        ins_ctx   = smgr.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", ctx)
        scriptPro = ins_ctx.createScriptProvider("")
        location  = "application"
    
    scriptNome = "vnd.sun.star.script:"+library+"."+module+"."+macro+ "?language=Basic&location="+location
    
    return scriptPro.getScript(scriptNome)

def LO_Atualizar():
    """
    Atualiza o LibreOffice para não travar/congelar.
    """
    basScript = LO_ScriptBasic()
    basScript.invoke((), (), ())
    return 0
    
def Saida(texto, log = '', saida = '', dlg = False):
    """
    Saida de 'texto' para acompanhamento.
    Pode acrescentar no log (tela do script ou objeto do LibreOffice),
    pode escrever em uma saida e pode gerar uma MsgBox do LibreOffice.
    """
    from datetime                     import datetime
    
    global LO
    
    msg = "{0} - {1}".format(datetime.now().strftime("%d/%m/%Y %H:%M:%S"), texto)
    
    if log:
        log.String += '\n' + msg
    else:
        print(msg)
    
    if saida:
        saida.String = msg
    
    if dlg:
        # Método antigo
        # janela = Sobre.MsgBox(XSCRIPTCONTEXT.getComponentContext())
        # janela.addButton("OK")
        # janela.renderFromButtonSize(75)
        # janela.numberOflines = 2
        # janela.show(texto,0,titulo)
        bas = CreateScriptService("Basic")
        bas.MsgBox(texto)
    
    if LO:
        LO_Atualizar()
    
    return 0

def Sobre(*args):
    texto = """
    Versão:   30/11/2023
    Autor:    CC (EN) Lauro Sá
    OM:       CGPIM/SGM
    Licença:  MIT
              (deve-se citar ou fazer referência a fonte e autor)
    """
    
    Saida(texto = texto, dlg = True)
    
    return 0

def LimpaRIP(RIP):
    return ''.join(i for i in RIP if i.isdigit())
    # return RIP.replace(' ', '').replace('.', '').replace('-', '')

def Sessao():
    """
    Sessão do WebDriver.
    """
    import os
    import platform
    
    options          = ChromeOptions()
    options.headless = True
    options.add_argument("--headless")
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--no-default-browser-check')
    options.add_argument('--no-first-run')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-default-apps')
    options.add_argument("--window-size=1024,768")
    options.add_argument("--ignore-certificate-errors")
    options.add_experimental_option("prefs", {
    "download.default_directory":         str(BASE_DIR()),
    "download.prompt_for_download":       False,
    "download.directory_upgrade":         True,
    "plugins.always_open_pdf_externally": True
    })
    if platform.system() == "Windows":
        options.binary_location = os.path.join(BASE_DIR(), 'chrome', 'chrome.exe')
        engine_path             = os.path.join(BASE_DIR(), 'chrome', 'chromedriver.exe')
    else:
        options.binary_location = os.path.join(BASE_DIR(), 'chrome', 'chrome')
        engine_path             = os.path.join(BASE_DIR(), 'chrome', 'chromedriver')
    service = Service(executable_path=engine_path)
    
    return Chrome(service=service, options=options)

class Raspador():
    """Raspador para SPIUnet"""
    from datetime                     import date
    
    URL_P = 'http://spiunet.spu.planejamento.gov.br/default.asp'
    URL_U = "http://spiunet.spu.planejamento.gov.br/consulta/Cons_Utilizacao.asp?NU_RIP={}"
    URL_I = "http://spiunet.spu.planejamento.gov.br/consulta/Cons_Imovel.asp?NU_RIP={}"
    CAMPOS  = {
        'RIPI': (str(), 'Rip:', 1),
        'IncorporaData': (date(2023, 11, 30), 'Data da Incorporação:', 1),
        'Logradouro': (str(), 'Logradouro:', 1),
        'Numero': (int(), 'Número:', 1),
        'Natureza': (str(), 'Natureza:', 1),
        'TerrenoAreaI': (float(), """Área
                  Terreno (m²):""", 1),
        'ConstruidaAreaI': (float(), 'Área Construída (m²):', 1),
        'TerrenoValorI': (float(), 'Valor do Terreno (R$):', 1),
        'BenfeitoriasValorI': (float(), 'Valor Benfeitorias Utilizações (R$):', 1),
        'ImovelValor': (float(), 'Valor do Imóvel (R$):', 1),
        
        'RIPU': (str(), 'RIP Utilização:', 1),
        'UGcod': (str(), 'Código UG/Gestão:', 2),
        'DestTipo': (str(), 'Tipo de Destinação:', 2),
        'DestDesc': (str(), 'Descrição da Destinação:', 1),
        'TerrenoAreaU': (float(), 'Área Terreno Utilizada (m²):', 1),
        'ConstruidaAreaU': (float(), 'Área Construída (m²):', 3),
         'Tipo': (str(), 'Tipo do Imóvel:', 1),
         'AvaliacaoDataU': (date(2023, 11, 30), 'Data Avaliação:', 1),
         'TerrenoValorU': (float(), 'Valor do Terreno (R$):', 1),
         'BenfeitoriasValorU': (float(), 'Valor da Benfeitoria (R$):', 1),
         'UtilizacaoValor': (float(), 'Valor da Utilização (R$):', 1)}
    ESPERA = '//body'
    PARSER = 'html.parser'
    HEADERS  = {
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'accept-encoding': 'gzip, deflate',
    'accept-language': 'pt,en-US;q=0.9,en;q=0.8',
    'cache-control': 'max-age=0',
    'connection': 'keep-alive',
    'content-length': '50',
    'content-type': 'application/x-www-form-urlencoded',
    'cookie': 'ASPSESSIONIDACTRCCBA=HAGDHOCAPKMEJDDECLCLCBBL',
    'host': 'spiunet.spu.planejamento.gov.br',
    'origin': 'http://spiunet.spu.planejamento.gov.br',
    'referer': 'http://spiunet.spu.planejamento.gov.br/Logon/Spiunet.asp',
    'upgrade-insecure-requests': '1',
    'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'
    }

    def __init__(self, usuario, senha, RIPs = '', campos = '*', saida = '', situacao = '', arquivo = '', headers = HEADERS, params = {}):
        self.usuario  = usuario
        self.senha    = senha
        self.saida    = saida
        self.situacao = situacao
        self.arquivo  = arquivo
        self.campos   = campos
        self.headers  = headers
        self.params   = params
        self.RIPs     = RIPs
        self.data     = []
        self.pages    = []
        self.timeout  = 3
        self.contador = 1
        self.engine   = Sessao()
    
    @property
    def campos(self):
        return self._campos

    @campos.setter
    def campos(self, value):
        self._campos = None
        if isinstance(value, list):
            self._campos = list(set(list(self.CAMPOS.keys())).intersection(set(value)))
        if value == '*':
            self._campos = list(self.CAMPOS.keys())
        if self._campos is None:
            Saida("{0} não é um objeto válido! Campos configurado como 'None'.".format(value), self.saida, self.situacao)

    def safe_text(self, obj, currency=False, strip = True):
        result = obj
        if obj:
            if isinstance(obj, str):
                result = obj
            else:
                result = obj.text
            if not currency:
                # TODO: implementar substituicao usando 'locale'
                result = result.replace('R$', '')
            if strip:
                result = result.strip()
        return result

    def login(self):
        Saida("Iniciando login no SPIUnet.", self.saida, self.situacao)
        self.engine.get(self.URL_P)
        self.engine.switch_to.frame("Principal")
        usuario = self.engine.find_element(By.NAME, "Login")
        senha   = self.engine.find_element(By.NAME, "Senha")
        if usuario and senha:
            usuario.send_keys(self.usuario)
            senha.send_keys(self.senha)
            self.engine.find_element(By.XPATH, "//input[@value='Avançar']").click()
            Saida("Finalizado login no SPIUnet.", self.saida, self.situacao)
        else:
            Saida("Usuário já logado no SPIUnet.", self.saida, self.situacao)
        
    def get_data(self, page):
        result = []
        dado_r = {}
        if page.find('font', string='Msg: 0017 - RIP não cadastrado.') or \
           page.find('font', string='500 - Internal server error.') or \
           page.find('404 - File or directory not found.'):
            for campo in self.campos:
                dado_r[campo] = ''
        else:
            for campo in self.campos:
                resultado  = ''
                resultados = page.findAll(text=self.CAMPOS[campo][1])
                for resultado in resultados:
                    try:
                        if self.CAMPOS[campo][2] == 1:
                            resultado = self.safe_text(resultado.find_parent().find_parent().findNextSibling().font.b.text)
                        if self.CAMPOS[campo][2] == 2:
                            resultado = self.safe_text(resultado.find_parent().find_parent().findNextSibling().b.font.text)
                        if self.CAMPOS[campo][2] == 3:
                            resultado = self.safe_text(resultado.find_parent().find_parent().find_parent().findNextSibling().font.b.text)
                    except:
                        pass
                dado_r[campo] = resultado
            
        result.append(dado_r)
        return result

    def get_pages(self):
        import csv
        from datetime                     import datetime
        
        self.login()
        dt_ini = datetime.now()
        Saida("Raspagem iniciada.", self.saida, self.situacao)
        self.pages = []
        self.data = []

        if self.arquivo:
            with open(self.arquivo, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames = ['RIP',]+self.campos)
                writer.writeheader()

        RIPs = self.RIPs
        if isinstance(RIPs, str):
            arquivo = open(RIPs, 'r')
            RIPs = csv.DictReader(arquivo, fieldnames=['RIP',])
        else:
            RIPs = []
            for RIP in self.RIPs:
                RIPs.append({'RIP': RIP})

        for RIP in RIPs:
            dt_pg = datetime.now()
            Saida("RIP {0} - Raspando.".format(RIP['RIP']), self.saida, self.situacao)
            urlN      = self.URL_U.format(RIP['RIP'])
            self.engine.get(urlN)
            try:
                next_page = WebDriverWait(self.engine, self.timeout).until(visibility_of_element_located((By.XPATH, self.ESPERA)))
            except TimeoutException:
                try:
                    urlN      = self.URL_I.format(RIP['RIP'])
                    self.engine.get(urlN)
                    next_page = WebDriverWait(self.engine, self.timeout).until(visibility_of_element_located((By.XPATH, self.ESPERA)))
                except TimeoutException:
                    Saida("RIP {0} - Não encontrado.".format(RIP['RIP']), self.saida, self.situacao)
                    next

            if self.engine.current_url == self.URL_P:
                self.engine.switch_to.frame("Principal")
            self.pages.append(BeautifulSoup(self.engine.page_source, self.PARSER))
            rows = self.get_data(self.pages[-1])
            rows[0]['RIP'] = str(RIP['RIP'])
            self.data.extend(rows)
            if self.arquivo:
                with open(self.arquivo, 'a', newline='', encoding='utf-8') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames = ['RIP',]+self.campos)
                    writer.writerows(rows)
            Saida("RIP {0} - Raspado. Tempo decorrido: {1}.".format(RIP['RIP'], (datetime.now() - dt_pg)), self.saida, self.situacao)

            self.contador += 1

        Saida("Raspagem finalizada. Tempo decorrido: {0}.".format((datetime.now() - dt_ini)), self.saida, self.situacao)
        return self.data
