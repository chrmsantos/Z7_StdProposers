import os
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

DEFAULT_PROMPT = """Você é um especialista em revisão de textos legislativos no idioma Português do Brasil.
Abaixo está um trecho de uma propositura legislativa.
Sua tarefa é corrigir gramaticalmente o texto, realizando O MÍNIMO POSSÍVEL de alterações em relação ao original.
Mantenha o tom formal, o jargão jurídico/legislativo e a estrutura da frase intactos, corrigindo apenas erros de ortografia, concordância, regência, pontuação ou crase evidentes.
Não adicione ponto final se o texto original não possuir um.
Retorne APENAS o texto corrigido, sem adicionar nenhum comentário, explicação, formatação markdown ou aspas extras."""

def get_prompt_file_path():
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        return None
    key_dir = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
    key_dir.mkdir(parents=True, exist_ok=True)
    return key_dir / 'gemini_prompt.txt'

def load_prompt():
    prompt_file = get_prompt_file_path()
    if prompt_file and prompt_file.exists():
        try:
            with open(prompt_file, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception:
            pass
    return DEFAULT_PROMPT

def save_prompt(text_widget, root):
    new_prompt = text_widget.get("1.0", tk.END).strip()
    if not new_prompt:
        messagebox.showwarning("Aviso", "O prompt não pode estar vazio.")
        return

    prompt_file = get_prompt_file_path()
    if prompt_file:
        try:
            with open(prompt_file, 'w', encoding='utf-8') as f:
                f.write(new_prompt)
            messagebox.showinfo("Sucesso", "Prompt configurado com sucesso!")
            root.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar o prompt:\n{e}")

def main():
    root = tk.Tk()
    root.title("Configurar Prompt do Gemini")
    root.geometry("600x450")
    
    # Faz a janela aparecer na frente
    root.attributes('-topmost', True)
    
    lbl = tk.Label(root, text="Instruções para a Inteligência Artificial:", font=("Arial", 12, "bold"))
    lbl.pack(pady=(15, 5))
    
    info_lbl = tk.Label(root, text="Personalização das instruções enviadas para a IA.", font=("Arial", 9, "italic"), fg="#555555")
    info_lbl.pack(pady=(0, 10))

    frame = tk.Frame(root)
    frame.pack(expand=True, fill=tk.BOTH, padx=20)

    text_area = tk.Text(frame, wrap=tk.WORD, font=("Consolas", 10), relief=tk.GROOVE, borderwidth=2)
    text_area.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

    scrollbar = tk.Scrollbar(frame, command=text_area.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    text_area.config(yscrollcommand=scrollbar.set)

    # Carrega o prompt atual
    current_prompt = load_prompt()
    text_area.insert(tk.END, current_prompt)

    btn_frame = tk.Frame(root)
    btn_frame.pack(fill=tk.X, pady=15)

    save_btn = tk.Button(btn_frame, text="Salvar Configuração", width=20, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), command=lambda: save_prompt(text_area, root))
    save_btn.pack(side=tk.RIGHT, padx=20)
    
    cancel_btn = tk.Button(btn_frame, text="Cancelar", width=15, font=("Arial", 10), command=root.destroy)
    cancel_btn.pack(side=tk.RIGHT, padx=5)

    root.mainloop()

if __name__ == "__main__":
    main()
