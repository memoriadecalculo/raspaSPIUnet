#!/usr/bin/env python3
# coding: utf-8

from datetime     import datetime
from raspaSPIUnet import Raspador, Perfil

# Dados do SPIUnet
usuario  = '...........'
senha    = '...........'

# Entrada
urls_csv = '.............csv'

# Resultados
arquivo  = 'resultado{0}.csv'.format(datetime.now().strftime("%Y%m%d-%H%M%S"))

# Raspagem dos dados
SPIUnetTeste   = Perfil(usuario=usuario, senha=senha, arquivo=arquivo, urls_csv=urls_csv)
SPIUnetScraper = Raspador(SPIUnetTeste)
SPIUnetScraper.get_pages()
