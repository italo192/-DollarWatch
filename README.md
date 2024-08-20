# DollarWatch

**DollarWatch** é um aplicativo automatizado que realiza a consulta diária da cotação do dólar em um site específico, captura um screenshot da página, gera um relatório em Word com as informações coletadas e converte o relatório para PDF. Este projeto foi desenvolvido com o objetivo de simplificar o monitoramento da cotação do dólar, automatizando o processo de coleta e registro das informações.

## Funcionalidades

- Acessa automaticamente o site da cotação do dólar.
- Coleta o valor da cotação e a data/hora da consulta.
- Captura um screenshot da página de cotação.
- Gera um relatório em Word com os dados coletados e a imagem capturada.
- Converte o relatório em PDF.
- Verifica se a cotação já foi atualizada no dia corrente para evitar duplicidade.
- Oferece feedback detalhado no terminal sobre cada etapa do processo.
- Executa a consulta diariamente de forma automática no horário agendado.
- Permite inicialização manual do processo através de uma tecla.

## Requisitos

- Python 3.x
- Selenium
- ChromeDriver
- python-docx
- docx2pdf
- schedule

## Instalação

1. Clone este repositório:

   ```bash
   git clone https://github.com/seu-usuario/dollarwatch.git
