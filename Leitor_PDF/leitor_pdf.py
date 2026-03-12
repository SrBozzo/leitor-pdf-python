import fitz  # A biblioteca PyMuPDF
import os

def extrair_texto_pdf(caminho_arquivo):
    texto_completo = ""
    
    if not os.path.exists(caminho_arquivo):
        return f"Erro: O arquivo '{caminho_arquivo}' não foi encontrado."

    try:
        # O 'with' é a mágica aqui: ele garante que o ficheiro é fechado e 
        # libertado da memória do Windows assim que a leitura terminar!
        with fitz.open(caminho_arquivo) as documento:
            for pagina in documento:
                # Extrai o texto preservando parágrafos básicos
                texto_completo += pagina.get_text("text") + "\n"
        
        # Verificação inteligente: Se o texto estiver vazio após ler todas as páginas
        if not texto_completo.strip():
            return (
                "⚠️ Aviso: Nenhum texto detetado neste PDF.\n\n"
                "Motivo provável: Este documento foi digitalizado como uma imagem (ex: documento no scanner) "
                "e não possui texto selecionável nativo."
            )
            
        return texto_completo
        
    except Exception as e:
        return f"❌ Ocorreu um erro ao ler o PDF: {e}"