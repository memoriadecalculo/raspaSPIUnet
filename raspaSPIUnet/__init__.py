#!/usr/bin/env python3
# coding: utf-8

import csv
import time
from bs4                               import BeautifulSoup
from datetime                          import datetime
from os.path                           import join
from pathlib                           import Path
from selenium                          import webdriver
from selenium.common.exceptions        import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by      import By
from selenium.webdriver.support.ui     import WebDriverWait
from selenium.webdriver.support        import expected_conditions
from urllib.parse                      import urlparse, urlunparse
import platform

BASE_DIR = Path(__file__).resolve().parent.parent

class Perfil():
    """Perfil do Raspador para SPIUnet.
    
    url: url inicial.
         Se mais de uma pagina for raspada, o parametro "{}"
         dentro da url deve substituido pelo numero da pagina.
    
    qtdPg: quantidade de paginas a raspar.
            Raspador comecara pela 1 e parara na qtdPg.
    
    params: parametros para a requisicao da web.
    
    headers: para a requisicao da web."""
    
    CAMPOS = ['RIPI', 'Logradouro', 'Numero', 'TerrenoAreaI', 'ConstruidaAreaI', 'RIPU', 'TerrenoAreaU', 'ConstruidaAreaU', 'Tipo', 'AvaliacaoData', 'TerrenoValor', 'BenfeitoriasValor']
    ESPERA = '//body'
    PARSERS = ['html.parser',]
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
    
    def __init__(self, usuario, senha, arquivo = '', campos = '*', qtdPg = 1, headers = HEADERS, params = {}, urls_csv = ''):
        self.url      = "http://spiunet.spu.planejamento.gov.br/consulta/Cons_Utilizacao.asp?NU_RIP={}"
        self.usuario  = usuario
        self.senha    = senha
        self.arquivo  = arquivo
        self.campos   = campos
        self.qtdPg    = qtdPg
        self.headers  = headers
        self.params   = params
        self.parser   = 'html.parser'
        self.derivado = ''
        self.tipo     = 'Imóvel'
        self.urls_csv = urls_csv
    
    @property
    def campos(self):
        return self._campos
    
    @campos.setter
    def campos(self, value):
        self._campos = None
        if isinstance(value, list):
            self._campos = list(set(self.CAMPOS).intersection(set(value)))
        if value == '*':
            self._campos = self.CAMPOS
        if self._campos is None:
            print("{0} não é um objeto válido! Campos configurado como 'None'.".format(value))
    
    @property
    def url(self):
        return self._url.geturl()
    
    @url.setter
    def url(self, value):
        if value:
            self._url     = urlparse(value)
            self.url_base = urlunparse(self._url[:2] + ('', '', None, None))
        
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
    
    def get_data(self, page, url_base):
        result = []
        dado_r = {}
        if page.find('font', string='Msg: 0017 - RIP não cadastrado.'):
            dado_r['RIPI'] = dado_r['Logradouro'] = dado_r['Numero'] = dado_r['TerrenoAreaI'] = dado_r['ConstruidaAreaI'] = dado_r['RIPU'] = dado_r['TerrenoAreaU'] = dado_r['ConstruidaAreaU'] = dado_r['Tipo'] = dado_r['AvaliacaoData'] = dado_r['TerrenoValor'] = dado_r['BenfeitoriasValor'] = ''
        else:
            if 'RIPI' in self.campos:
                RIPI = page.find(text='Rip:')
                if RIPI is None:
                    RIPI = ''
                else:
                    RIPI = self.safe_text(RIPI.find_parent().find_parent().findNextSibling().font.b.text)
                dado_r['RIPI'] = RIPI
            
            if 'Logradouro' in self.campos:
                logradouro = page.find(text='Logradouro:')
                if logradouro is None:
                    logradouro = ''
                else:
                    logradouro = self.safe_text(logradouro.find_parent().find_parent().findNextSibling().font.b.text)
                dado_r['Logradouro'] = logradouro
            
            if 'Numero' in self.campos:
                numero = page.find(text='Número:')
                if numero is None:
                    numero = ''
                else:
                    numero = self.safe_text(numero.find_parent().find_parent().findNextSibling().font.b.text)
                dado_r['Numero'] = numero
            
            if 'TerrenoAreaI' in self.campos:
                TerrenoAreaI = page.find(text="""Área
                  Terreno (m²):""")
                if TerrenoAreaI is None:
                    TerrenoAreaI = ''
                else:
                    TerrenoAreaI = self.safe_text(TerrenoAreaI.find_parent().find_parent().findNextSibling().font.b.text)
                dado_r['TerrenoAreaI'] = TerrenoAreaI
            
            if ('ConstruidaAreaI' or 'ConstruidaAreaU') in self.campos:
                ConstruidaAreas = page.findAll(text='Área Construída (m²):')
                ConstruidaAreaI = ConstruidaAreaU = ''
                for area in ConstruidaAreas:
                    try:
                        ConstruidaAreaI = self.safe_text(area.find_parent().find_parent().findNextSibling().font.b.text)
                    except:
                        ConstruidaAreaU = self.safe_text(area.find_parent().find_parent().find_parent().findNextSibling().font.b.text)
                
                if 'ConstruidaAreaI' in self.campos:
                    dado_r['ConstruidaAreaI'] = ConstruidaAreaI
                
                if 'ConstruidaAreaU' in self.campos:
                    dado_r['ConstruidaAreaU'] = ConstruidaAreaU
                
            if 'RIPU' in self.campos:
                RIPU = page.find(text='RIP Utilização:')
                if RIPU is None:
                    RIPU = ''
                else:
                    RIPU = self.safe_text(RIPU.find_parent().find_parent().findNextSibling().font.b.text)
                dado_r['RIPU'] = RIPU
            
            if 'TerrenoAreaU' in self.campos:
                TerrenoAreaU = page.find(text='Área Terreno Utilizada (m²):')
                if TerrenoAreaU is None:
                    TerrenoAreaU = ''
                else:
                    TerrenoAreaU = self.safe_text(TerrenoAreaU.find_parent().find_parent().findNextSibling().font.b.text)
                dado_r['TerrenoAreaU'] = TerrenoAreaU
            
            if 'Tipo' in self.campos:
                tipo = page.find(text='Tipo do Imóvel:')
                if tipo is None:
                    tipo = ''
                else:
                    tipo = self.safe_text(tipo.find_parent().find_parent().findNextSibling().font.b.text)
                dado_r['Tipo'] = tipo
            
            if 'AvaliacaoData' in self.campos:
                data = page.find(text='Data Avaliação:')
                if data is None:
                    data = ''
                else:
                    data = self.safe_text(data.find_parent().find_parent().findNextSibling().font.b.text)
                dado_r['AvaliacaoData'] = data
            
            if 'TerrenoValor' in self.campos:
                terreno = page.find(text='Valor do Terreno (R$):')
                if terreno is None:
                    terreno = ''
                else:
                    terreno = self.safe_text(terreno.find_parent().find_parent().findNextSibling().font.b.text)
                dado_r['TerrenoValor'] = terreno
            
            if 'BenfeitoriasValor' in self.campos:
                benfeitoria = page.find(text='Valor Benfeitorias Utilizações (R$):')
                if benfeitoria is None:
                    benfeitoria = ''
                else:
                    benfeitoria = self.safe_text(benfeitoria.find_parent().find_parent().findNextSibling().font.b.text)
                dado_r['BenfeitoriasValor'] = benfeitoria
            
        result.append(dado_r)
        return result
    
    def get_all_data(self, pages):
        result = []
        return result

class Raspador():
    """Raspador para SPIUnet"""
    
    SPIUnet_URL_BASE = 'http://spiunet.spu.planejamento.gov.br/default.asp'
    
    def __init__(self, profile):
        self.profile          = profile
        self.data             = []
        self.pages            = []
        self.timeout          = 3
        self.contador         = 1
        self.options          = Options()
        self.options.headless = True
        self.options.add_argument('--no-sandbox')
        self.options.add_argument('--disable-dev-shm-usage')
        self.options.add_argument('--no-default-browser-check')
        self.options.add_argument('--no-first-run')
        self.options.add_argument('--disable-gpu')
        self.options.add_argument('--disable-extensions')
        self.options.add_argument('--disable-default-apps')
        self.options.add_argument("--window-size=1024,768")
        if platform.system() == "Windows":
            self.options.binary_location = join(BASE_DIR, 'GoogleChrome', 'App', 'Chrome-bin', 'chrome.exe')
            self.engine_path             = join(BASE_DIR, 'GoogleChromeDriver', 'chromedriver.exe')
        else:
            self.options.binary_location = '/opt/google/chrome/chrome'
            self.engine_path             = join(BASE_DIR, 'GoogleChromeDriver', 'chromedriver')
        
        self.engine           = webdriver.Chrome(options=self.options, executable_path=self.engine_path)
        print("{0} - Iniciando login no SPIUnet.".format(datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
        self.engine.get(self.SPIUnet_URL_BASE)
        self.engine.switch_to.frame("Principal")
        usuario = self.engine.find_element(By.NAME, "Login")
        senha   = self.engine.find_element(By.NAME, "Senha")
        usuario.send_keys(self.profile.usuario)
        senha.send_keys(self.profile.senha)
        self.engine.find_element(By.XPATH, "//input[@value='Avançar']").click()
        print("{0} - Finalizado login no SPIUnet.".format(datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
    
    @property
    def profile(self):
        return self._profile
    
    @profile.setter
    def profile(self, value):
        self._profile = None
        if value:
            if isinstance(value, Perfil):
                self._profile = value
            else:
                print("{0} não é um objeto Perfil! Salvando 'None' como perfil.".format(value))
        else:
            print("Salvando 'None' como perfil.".format(value))
    
    def get_pages(self):
        dt_ini = datetime.now()
        print("{0} - {1} - Raspagem iniciada.".format(dt_ini.strftime("%d/%m/%Y %H:%M:%S"), self.profile.tipo))
        self.pages = []
        self.data = []
        
        if self.profile.arquivo:
            with open(self.profile.arquivo, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames = ['RIP',]+self.profile.campos)
                writer.writeheader()
        
        arquivo = open(self.profile.urls_csv, 'r')
        urls_csv = csv.DictReader(arquivo, fieldnames=['RIP',])
        for url in urls_csv:
            dt_pg = datetime.now()
            print("{0} - {1} {2} - Raspando.".format(dt_pg.strftime("%d/%m/%Y %H:%M:%S"), self.profile.tipo, url['RIP']))
            urlN      = self.profile.url.format(url['RIP'])
            print(urlN)
            self.engine.get(urlN)
            print(self.engine.current_url)
            try:
                next_page = WebDriverWait(self.engine, self.timeout).until(expected_conditions.visibility_of_element_located((By.XPATH, self.profile.ESPERA)))
            except TimeoutException:
                try:
                    urlN      = "http://spiunet.spu.planejamento.gov.br/consulta/Cons_Imovel.asp?NU_RIP={}".format(url['RIP'])
                    self.engine.get(urlN)
                    print(self.engine.current_url)
                    next_page = WebDriverWait(self.engine, self.timeout).until(expected_conditions.visibility_of_element_located((By.XPATH, self.profile.ESPERA)))
                except TimeoutException:
                    print("{0} - {1} {2} - Não encontrado.".format(datetime.now().strftime("%d/%m/%Y %H:%M:%S"), self.profile.tipo, url['RIP']))
                    next
            
            if self.engine.current_url == self.SPIUnet_URL_BASE:
                self.engine.switch_to.frame("Principal")
                # print('FRAME')
            self.pages.append(BeautifulSoup(self.engine.page_source, self.profile.parser))
            # TODO
            # teste = self.profile.get_data(self.pages[-1], self.profile.url_base)[0]
            # teste = ['RIP'] = url['RIP']
            # print(teste)
            rows = self.profile.get_data(self.pages[-1], self.profile.url_base)
            rows[0]['RIP'] = str(url['RIP'])
            # print(rows)
            self.data.extend(rows)
            if self.profile.arquivo:
                with open(self.profile.arquivo, 'a', newline='', encoding='utf-8') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames = ['RIP',]+self.profile.campos)
                    writer.writerows(rows)
            print("{0} - {1} {2} - Raspado. Tempo decorrido: {3}.".format(datetime.now().strftime("%d/%m/%Y %H:%M:%S"), self.profile.tipo, url['RIP'], (datetime.now() - dt_pg)))
            
            self.contador += 1
            
        print("{0} - {1} - Raspagem finalizada. Tempo decorrido: {2}.".format(datetime.now().strftime("%d/%m/%Y %H:%M:%S"), self.profile.tipo, (datetime.now() - dt_ini)))
        return self.data
    
    def get_data(self):
        self.data = self.profile.get_data(self.pages, self.profile.url_base)
    
    def save_header(self):
        with open(self.profile.arquivo, 'w') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames = self.profile.campos)
            writer.writeheader()
    
    def save_rows(self, rows):
        with open(self.profile.arquivo, 'w') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames = self.profile.campos)
            writer.writerows(rows)
