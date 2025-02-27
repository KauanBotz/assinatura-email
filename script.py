import os
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import win32com.client

def editar_assinatura(pptx_path, nome, funcao, numero=None, endereco=None):
    if not os.path.exists(pptx_path):
        print(f"Erro: O arquivo '{pptx_path}' não foi encontrado!")
        return
    
    print(f"Arquivo '{pptx_path}' carregado com sucesso!")

    prs = Presentation(pptx_path)
    
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                if '{{Nome}}' in shape.text:
                    shape.text = shape.text.replace('{{Nome}}', nome)
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)  # Cor #595959
                if '{{Funcao}}' in shape.text:
                    shape.text = shape.text.replace('{{Funcao}}', funcao)
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)  # Cor #595959
                if '{{Numero}}' in shape.text and numero:
                    shape.text = shape.text.replace('{{Numero}}', numero)
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(0x3A, 0x8C, 0x62)  # Cor #3A8C62
                elif '{{Numero}}' in shape.text:
                    shape.text = shape.text.replace('{{Numero}}', '')
                if '{{Endereço}}' in shape.text and endereco:
                    shape.text = shape.text.replace('{{Endereço}}', endereco)
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(0x3A, 0x8C, 0x62)  # Cor #3A8C62
                elif '{{Endereço}}' in shape.text:
                    shape.text = shape.text.replace('{{Endereço}}', '')

    temp_pptx_path = 'temp_assinatura.pptx'
    prs.save(temp_pptx_path)  
    print(f"PPTX modificado salvo como '{temp_pptx_path}'.")
    
    jpg_path = pptx_path.replace('.pptx', f'_{nome}.jpg')
    converter_para_jpg(temp_pptx_path, jpg_path)

def converter_para_jpg(pptx_path, jpg_path):
    ppt = None
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.Visible = True 
        presentation = ppt.Presentations.Open(os.path.abspath(pptx_path))

        presentation.SaveAs(os.path.abspath(jpg_path), 18)
        presentation.Close()
        print(f"Arquivo JPG gerado e salvo como '{jpg_path}'.")

    except Exception as e:
        print(f"Erro ao converter para JPG: {e}")
    finally:
        if ppt:
            ppt.Quit()

def escolher_endereco():
    print("Escolha o endereço:")
    print("(1) Rua São Pedro da Aldeia, 1200 - Pilar")
    print("(2) Governador Valadares")
    print("(3) Conselheiro Pena")
    escolha = input("Digite o número da opção desejada: ")
    
    enderecos = {
        "1": "Rua São Pedro da Aldeia, 1200 - Pilar",
        "2": "Governador Valadares",
        "3": "Conselheiro Pena"
    }
    
    return enderecos.get(escolha, "Endereço não especificado")

nome = input("Digite o nome: ")
funcao = input("Digite a função: ")
numero = input("Digite o seu número de telefone: ")
endereco = escolher_endereco()

editar_assinatura(
    pptx_path='modelo-assinaturaNovo.pptx',
    nome=nome,
    funcao=funcao,
    numero=numero if numero else None,
    endereco=endereco
)