# Projeto de Automação de Assinaturas de E-mail

Este projeto tem como objetivo automatizar a criação de assinaturas de e-mail personalizadas para os funcionários de uma empresa. Utilizando **Python**, o código modifica um arquivo **PPTX** (modelo de assinatura) com base nas informações de cada usuário e gera a assinatura personalizada em formato **JPG**.

## Funcionalidades

- **Substituição dinâmica** de placeholders (`{{Nome}}`, `{{Funcao}}`, `{{Numero}}`, `{{Endereço}}`) por informações reais.
- Geração automática de assinaturas em **PPTX** e conversão para **JPG**.
- Personalização de cores de texto de acordo com a categoria de informação.
- Automação de um processo manual de criação de assinaturas, tornando-o mais rápido e eficiente.

## Tecnologias Usadas

- **Python** com bibliotecas `python-pptx` para manipulação do PPTX e `win32com.client` para conversão para imagens.
- **PowerPoint** (instalado no computador) para conversão de PPTX para JPG.

## Como Usar

1. Tenha o PowerPoint instalado no seu computador.
2. Instale as dependências necessárias:
   ```bash
   pip install python-pptx pywin32
