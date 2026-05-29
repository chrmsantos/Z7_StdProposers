import os
import shutil
import sys
from pathlib import Path

def main():
    # Encontra o diretório da execução
    if getattr(sys, 'frozen', False):
        # Se executado como binário compilado pelo PyInstaller
        current_dir = Path(sys.executable).parent
    else:
        # Se executado como script Python
        current_dir = Path(__file__).resolve().parent

    # Identifica o diretório raiz do projeto (um nível acima de 'scripts')
    if current_dir.name.lower() == 'scripts':
        project_root = current_dir.parent
    else:
        project_root = Path(os.getcwd())
        if not (project_root / 'dist').exists():
            project_root = current_dir.parent

    dist_dir = project_root / 'dist'
    
    # Arquivos de origem
    src_normal = dist_dir / 'Normal.dotm'
    src_officeui = dist_dir / 'Word.officeUI'

    # Verifica se os arquivos de origem existem em dist
    if not src_normal.exists():
        print(f"Erro: Arquivo de origem 'Normal.dotm' nao encontrado em: {src_normal}")
        sys.exit(1)

    if not src_officeui.exists():
        print(f"Erro: Arquivo de origem 'Word.officeUI' nao encontrado em: {src_officeui}")
        sys.exit(1)

    # Caminho de destino para o Normal.dotm
    appdata = os.environ.get('APPDATA')
    userprofile = os.environ.get('USERPROFILE')

    dest_normal_dir = None
    if appdata:
        dest_normal_dir = Path(appdata) / 'Microsoft' / 'Templates'
    elif userprofile:
        dest_normal_dir = Path(userprofile) / 'AppData' / 'Roaming' / 'Microsoft' / 'Templates'

    if not dest_normal_dir:
        print("Erro: Nao foi possivel determinar o diretorio AppData de destino.")
        sys.exit(1)

    # Caminho de destino para o Word.officeUI
    localappdata = os.environ.get('LOCALAPPDATA')
    dest_officeui_dir = None
    if localappdata:
        dest_officeui_dir = Path(localappdata) / 'Microsoft' / 'Office'
    elif userprofile:
        dest_officeui_dir = Path(userprofile) / 'AppData' / 'Local' / 'Microsoft' / 'Office'

    if not dest_officeui_dir:
        print("Erro: Nao foi possivel determinar o diretorio LocalAppData de destino.")
        sys.exit(1)

    # Garante que os diretórios de destino existam
    dest_normal_dir.mkdir(parents=True, exist_ok=True)
    dest_officeui_dir.mkdir(parents=True, exist_ok=True)

    dest_normal = dest_normal_dir / 'Normal.dotm'
    dest_officeui = dest_officeui_dir / 'Word.officeUI'

    # Copia o Normal.dotm
    try:
        shutil.copy2(src_normal, dest_normal)
        print(f"Sucesso: 'Normal.dotm' importado para: {dest_normal}")
    except Exception as e:
        print(f"Erro ao importar 'Normal.dotm': {e}")
        if isinstance(e, PermissionError) or (hasattr(e, 'winerror') and getattr(e, 'winerror') == 32):
            print("\n>>> Dica: Certifique-se de fechar o Microsoft Word antes de importar o Normal.dotm. <<<")
        sys.exit(1)

    # Copia o Word.officeUI
    try:
        shutil.copy2(src_officeui, dest_officeui)
        print(f"Sucesso: 'Word.officeUI' importado para: {dest_officeui}")
    except Exception as e:
        print(f"Erro ao importar 'Word.officeUI': {e}")
        if isinstance(e, PermissionError) or (hasattr(e, 'winerror') and getattr(e, 'winerror') == 32):
            print("\n>>> Dica: Certifique-se de fechar o Microsoft Word antes de importar o Word.officeUI. <<<")
        sys.exit(1)

if __name__ == '__main__':
    main()
