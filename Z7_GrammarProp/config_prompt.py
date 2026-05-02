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

def restore_default(text_widget):
    text_widget.delete("1.0", tk.END)
    text_widget.insert(tk.END, DEFAULT_PROMPT)

def main():
    root = tk.Tk()
    root.title("Configurar Prompt do Gemini")
    root.geometry("700x600")
    root.minsize(600, 500)
    root.configure(bg="#f3f4f6")
    
    # Faz a janela aparecer na frente
    root.attributes('-topmost', True)
    
    lbl = tk.Label(root, text="Instruções para a Inteligência Artificial", font=("Segoe UI", 16, "bold"), bg="#f3f4f6", fg="#111827")
    lbl.pack(pady=(25, 5))
    
    info_lbl = tk.Label(root, text="Personalize o comportamento do modelo ajustando o prompt abaixo.", font=("Segoe UI", 10), bg="#f3f4f6", fg="#4b5563")
    info_lbl.pack(pady=(0, 20))

    btn_frame = tk.Frame(root, bg="#f3f4f6")

    frame = tk.Frame(root, bg="#d1d5db") # Borda sutil usando cor de fundo do frame
    text_area = tk.Text(frame, wrap=tk.WORD, font=("Consolas", 11), bg="#ffffff", fg="#1f2937", relief=tk.FLAT, padx=12, pady=12, insertbackground="#1f2937")
    text_area.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=1, pady=1)

    scrollbar = tk.Scrollbar(frame, command=text_area.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    text_area.config(yscrollcommand=scrollbar.set)

    # Carrega o prompt atual
    current_prompt = load_prompt()
    text_area.insert(tk.END, current_prompt)

    # Estilos de botão
    btn_font = ("Segoe UI", 10, "bold")
    
    save_btn = tk.Button(btn_frame, text="Salvar Configuração", width=20, bg="#2563eb", fg="white", font=btn_font, relief=tk.FLAT, activebackground="#1d4ed8", activeforeground="white", cursor="hand2", command=lambda: save_prompt(text_area, root))
    save_btn.pack(side=tk.RIGHT, padx=25)
    
    cancel_btn = tk.Button(btn_frame, text="Cancelar", width=15, bg="#e5e7eb", fg="#374151", font=btn_font, relief=tk.FLAT, activebackground="#d1d5db", activeforeground="#111827", cursor="hand2", command=root.destroy)
    cancel_btn.pack(side=tk.RIGHT, padx=5)

    restore_btn = tk.Button(btn_frame, text="Restaurar Padrão", width=18, bg="#e5e7eb", fg="#374151", font=btn_font, relief=tk.FLAT, activebackground="#d1d5db", activeforeground="#111827", cursor="hand2", command=lambda: restore_default(text_area))
    restore_btn.pack(side=tk.LEFT, padx=25)

    # Pack in order so btn_frame is fixed at the bottom
    btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=25)
    frame.pack(side=tk.TOP, expand=True, fill=tk.BOTH, padx=25, pady=(0, 10))

    root.mainloop()

if __name__ == "__main__":
    main()
