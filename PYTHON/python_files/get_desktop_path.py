import os
import ctypes

def get_desktop_path():
    """
    Obtém o caminho da área de trabalho do usuário atual.

    Retorna:
        str: Caminho da área de trabalho.
    """
    buffer = ctypes.create_string_buffer(1024)
    ctypes.windll.shell32.SHGetSpecialFolderPathW(None, ctypes.byref(buffer), CSIDL_DESKTOP, 0)
    return buffer.value.decode()

desktop_path = get_desktop_path()
print(f"Caminho da área de trabalho: {desktop_path}")