# Gerador de Assinaturas ✍️

Este é um programa em Python para gerar assinaturas personalizadas a partir de um modelo de PowerPoint (.pptx). O programa substitui placeholders no modelo (como {{Nome}}, {{Funcao}}, etc.) pelos dados fornecidos pelo usuário e salva a assinatura como uma imagem (.jpg).

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
