import os
import win32com.client
import google.generativeai as genai
import win32crypt
import tkinter as tk
from tkinter import simpledialog
from pathlib import Path

def get_api_key():
    # Define o caminho do arquivo de chave
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        print("Erro: Variavel USERPROFILE nao encontrada.")
        return None
        
    key_dir = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
    key_file = key_dir / 'gemini.key'
    
    # Se o arquivo ja existir, tenta ler e descriptografar
    if key_file.exists():
        try:
            with open(key_file, 'rb') as f:
                encrypted_key = f.read()
            # CryptUnprotectData retorna uma tupla (descricao, dados_descriptografados)
            _, decrypted_key = win32crypt.CryptUnprotectData(encrypted_key, None, None, None, 0)
            return decrypted_key.decode('utf-8')
        except Exception as e:
            print(f"Erro ao descriptografar a chave: {e}")
            # Se falhar, vamos pedir novamente
            pass
            
    # Se nao existir ou falhar, pede a chave via Tkinter
    root = tk.Tk()
    root.withdraw() # Esconde a janela principal
    
    # Mantem a janela de dialogo no topo
    root.attributes('-topmost', True)
    
    api_key = simpledialog.askstring(
        "Z7 StdProposers - Primeira Inicializacao", 
        "Insira sua chave da API do Google Gemini:\n(Ela sera criptografada e salva localmente)",
        parent=root
    )
    
    root.destroy()
    
    if not api_key or not api_key.strip():
        print("Chave nao fornecida pelo usuario.")
        return None
        
    api_key = api_key.strip()
    
    # Criptografa a chave usando DPAPI (seguranca atrelada ao usuario atual do Windows)
    try:
        encrypted_key = win32crypt.CryptProtectData(api_key.encode('utf-8'), 'Z7_Gemini_Key', None, None, None, 0)
        
        # Cria os diretorios se nao existirem
        key_dir.mkdir(parents=True, exist_ok=True)
        
        # Salva o arquivo
        with open(key_file, 'wb') as f:
            f.write(encrypted_key)
    except Exception as e:
        print(f"Erro ao criptografar ou salvar a chave: {e}")
        # Mesmo se falhar ao salvar, podemos retornar a chave em memoria para uso atual
        
    return api_key

def main():
    api_key = get_api_key()
    if not api_key:
        return

    # Configura a API do Gemini
    genai.configure(api_key=api_key)
    
    # Inicializa o modelo (gemini-1.5-pro e otimo para tarefas complexas de raciocinio, 
    # ou gemini-1.5-flash para respostas mais rapidas)
    model = genai.GenerativeModel('gemini-3.1-pro-preview')

    try:
        # Conecta ao aplicativo Word que já está em execução
        word = win32com.client.Dispatch("Word.Application")
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Z7 StdProposers - Erro", f"Erro ao conectar ao Word:\n{e}")
        return

    # Pega o texto selecionado pelo usuário
    selection = word.Selection
    text = selection.Text.strip()

    # Se nada estiver selecionado ou houver apenas espaços em branco, encerra
    if not text or len(text) < 2:
        return

    # Lê o prompt configurado do arquivo, se existir
    user_profile = os.environ.get('USERPROFILE')
    base_prompt = ""
    if user_profile:
        prompt_file = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers' / 'gemini_prompt.txt'
        if prompt_file.exists():
            try:
                with open(prompt_file, 'r', encoding='utf-8') as f:
                    base_prompt = f.read().strip()
            except Exception:
                pass
                
    if not base_prompt:
        base_prompt = """Você é um especialista em revisão de textos legislativos no idioma Português do Brasil.
Abaixo está um trecho de uma propositura legislativa.
Sua tarefa é corrigir gramaticalmente o texto, realizando O MÍNIMO POSSÍVEL de alterações em relação ao original.
Mantenha o tom formal, o jargão jurídico/legislativo e a estrutura da frase intactos, corrigindo apenas erros de ortografia, concordância, regência, pontuação ou crase evidentes.
Não adicione ponto final se o texto original não possuir um.
Retorne APENAS o texto corrigido, sem adicionar nenhum comentário, explicação, formatação markdown ou aspas extras."""

    prompt = f"{base_prompt}\n\nTexto original:\n{text}\n"

    try:
        # Faz a requisição para a API do Gemini
        response = model.generate_content(prompt)
        corrected_text = response.text.strip()
        
        if corrected_text:
            # Preserva a formatação substituindo o texto da seleção
            selection.Text = corrected_text
        else:
            root = tk.Tk()
            root.withdraw()
            messagebox.showwarning("Z7 StdProposers - Aviso", "A API do Gemini retornou um texto vazio.")
            
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Z7 StdProposers - Erro na API", f"Erro ao chamar a API do Gemini:\n{e}")

if __name__ == "__main__":
    main()
