"""
import_bas_to_normal.py
Imports .bas VBA module files from source/main/ into Normal.dotm's VBProject.
Requires: Word closed before running, or Word opened programmatically.
Uses win32com to interact with Word's VBA project.
"""
import os
import sys
import time
import glob
import tempfile
from pathlib import Path

def main():
    try:
        import win32com.client
    except ImportError:
        print("Erro: pywin32 nao instalado. Execute: pip install pywin32")
        sys.exit(1)

    # Resolve paths
    project_root = Path(__file__).resolve().parent.parent
    source_dir = project_root / "source" / "main"
    dist_dir = project_root / "dist"
    normal_dotm = dist_dir / "Normal.dotm"

    if not normal_dotm.exists():
        print(f"Erro: Normal.dotm nao encontrado em: {normal_dotm}")
        sys.exit(1)

    # Find all .bas files
    bas_files = sorted(glob.glob(str(source_dir / "*.bas")))
    if not bas_files:
        print(f"Erro: Nenhum arquivo .bas encontrado em: {source_dir}")
        sys.exit(1)

    print(f"Encontrados {len(bas_files)} arquivo(s) .bas:")
    for f in bas_files:
        print(f"  - {Path(f).name}")

    # Launch Word and open Normal.dotm
    print(f"\nAbrindo Normal.dotm: {normal_dotm}")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone

    try:
        doc = word.Documents.Open(str(normal_dotm))
        vb_project = doc.VBProject

        print(f"VBProject: {vb_project.Name}")
        print(f"Modulos existentes: {vb_project.VBComponents.Count}")

        # Import each .bas file
        imported_count = 0
        skipped_count = 0
        error_count = 0

        for bas_path in bas_files:
            bas_name = Path(bas_path).stem  # e.g., "Mod3Pipeline"
            # Source .bas files are UTF-8; try UTF-8 first, fall back to CP1252.
            raw = Path(bas_path).read_bytes()
            try:
                bas_content = raw.decode("utf-8")
            except (UnicodeDecodeError, ValueError):
                bas_content = raw.decode("cp1252")

            # Check if module already exists
            existing_module = None
            for comp in vb_project.VBComponents:
                if comp.Name == bas_name:
                    existing_module = comp
                    break

            try:
                if existing_module:
                    # Remove the existing module before re-importing.
                    # CodeModule.AddFromString() does NOT process Attribute VB_Name
                    # directives (it treats them as code, causing compile errors and
                    # leaving the module unnamed). VBComponents.Import() handles
                    # attributes correctly, so we always use Import for both new and
                    # existing modules.
                    vb_project.VBComponents.Remove(existing_module)

                # VBComponents.Import() reads the file using the system ANSI
                # codepage (CP1252 on Western European Windows). Write a temp
                # file in CP1252 so accented characters survive the round-trip.
                # Use errors='replace' to handle any character that cannot be
                # represented in CP1252 (e.g. U+FFFD replacement characters
                # or rare Unicode glyphs that crept into comments).
                tmp_path = None
                try:
                    fd, tmp_path = tempfile.mkstemp(suffix=".bas")
                    os.close(fd)
                    Path(tmp_path).write_text(bas_content, encoding="cp1252", errors="replace")
                    vb_project.VBComponents.Import(tmp_path)
                finally:
                    if tmp_path and os.path.exists(tmp_path):
                        try:
                            os.unlink(tmp_path)
                        except OSError:
                            pass

                imported_count += 1
                action = "re-importado" if existing_module else "importado como novo modulo"
                print(f"  [OK] {bas_name} - {action}")
            except Exception as e:
                error_count += 1
                print(f"  [ERRO] {bas_name} - {e}")

        # Save the document
        doc.Save()
        print(f"\nNormal.dotm salvo com sucesso!")
        print(f"Resumo: {imported_count} importados, {skipped_count} ignorados, {error_count} erros")

    except Exception as e:
        print(f"Erro durante processamento: {e}")
        error_count += 1
    finally:
        # Close Word
        try:
            doc.Close()
        except:
            pass
        try:
            word.Quit()
        except:
            pass

    if error_count > 0:
        print("\nATENCAO: Houve erros durante a importacao!")
        sys.exit(1)
    else:
        print("\nImportacao concluida com sucesso!")

if __name__ == "__main__":
    main()