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
    dist_dir.mkdir(parents=True, exist_ok=True)

    # Resolve os possíveis caminhos do Normal.dotm
    appdata = os.environ.get('APPDATA')
    userprofile = os.environ.get('USERPROFILE')

    possible_paths = []
    if appdata:
        appdata_path = Path(appdata)
        # Caminho padrão: AppData\Roaming\Microsoft\Templates\Normal.dotm
        possible_paths.append(appdata_path / 'Microsoft' / 'Templates' / 'Normal.dotm')
        # Caminho literal solicitado: AppData\Roaming\Roaming\Microsoft\Templates\Normal.dotm (por garantia)
        possible_paths.append(appdata_path / 'Roaming' / 'Microsoft' / 'Templates' / 'Normal.dotm')
    
    if userprofile:
        userprofile_path = Path(userprofile)
        possible_paths.append(userprofile_path / 'AppData' / 'Roaming' / 'Microsoft' / 'Templates' / 'Normal.dotm')

    # Busca pelo primeiro arquivo existente
    source_file = None
    for path in possible_paths:
        if path.exists():
            source_file = path
            break

    if not source_file:
        print("Erro: Arquivo 'Normal.dotm' nao encontrado nos seguintes caminhos:")
        for path in possible_paths:
            print(f"  - {path}")
        sys.exit(1)

    # Resolve os possíveis caminhos do Word.officeUI
    localappdata = os.environ.get('LOCALAPPDATA')
    possible_officeui_paths = []
    if localappdata:
        possible_officeui_paths.append(Path(localappdata) / 'Microsoft' / 'Office' / 'Word.officeUI')
    if userprofile:
        possible_officeui_paths.append(Path(userprofile) / 'AppData' / 'Local' / 'Microsoft' / 'Office' / 'Word.officeUI')

    source_officeui = None
    for path in possible_officeui_paths:
        if path.exists():
            source_officeui = path
            break

    if not source_officeui:
        print("Erro: Arquivo 'Word.officeUI' nao encontrado nos seguintes caminhos:")
        for path in possible_officeui_paths:
            print(f"  - {path}")
        sys.exit(1)

    # Cópia do Normal.dotm
    dest_file = dist_dir / 'Normal.dotm'
    try:
        shutil.copy2(source_file, dest_file)
        print(f"Sucesso: 'Normal.dotm' exportado para: {dest_file}")
    except Exception as e:
        print(f"Erro ao copiar 'Normal.dotm': {e}")
        sys.exit(1)

    # Cópia do Word.officeUI
    dest_officeui = dist_dir / 'Word.officeUI'
    try:
        shutil.copy2(source_officeui, dest_officeui)
        print(f"Sucesso: 'Word.officeUI' exportado para: {dest_officeui}")
    except Exception as e:
        print(f"Erro ao copiar 'Word.officeUI': {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
