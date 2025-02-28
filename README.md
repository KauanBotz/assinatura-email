# Projeto de Automação de Assinaturas de E-mail (Construtora Sant'Anna) ✍️

Este projeto tem como objetivo automatizar a criação de assinaturas de e-mail personalizadas para os funcionários da empresa em que trabalho (Construtora Sant'Anna). Utilizando **Python**, o código modifica um arquivo **PPTX** (modelo de assinatura) com base nas informações de cada usuário e gera a assinatura personalizada em formato **JPG**.

## Funcionalidades ⚙️

- **Substituição de Placeholders:**
  - Substitui {{Nome}}, {{Funcao}}, {{Numero}} e {{Endereco}} pelos dados fornecidos pelo usuário.
  
- **Validação de Número de Telefone:**
  - Valida se o número de telefone está no formato (XX) XXXX-XXXX ou (XX) 9XXXX-XXXX.
  
- **Geração de Várias Assinaturas:**
  - Permite gerar várias assinaturas sem reiniciar o programa.
  
- **Salvar em Subpasta:**
  - Salva os arquivos gerados em uma subpasta chamada `assinatura`.
  
- **Endereço Personalizado:**
  - Permite ao usuário digitar um endereço personalizado.
  
- **Conversão para JPG:**
  - Converte o arquivo .pptx modificado para .jpg.

## Requisitos 💻

- Python 3.x

### Bibliotecas:
- `python-pptx`
- `pywin32` (para usar o `win32com.client`)

- Microsoft PowerPoint instalado (para conversão de .pptx para .jpg).

## Melhorias Futuras 🔧

- Linkar a um Banco SQL
- Adicionar assinatura automaticamente no email de todos os usuários do @gruposantanna.com.br
