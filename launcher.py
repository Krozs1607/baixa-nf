"""
Launcher do executavel - configura paths e inicia o servidor.
Quando rodando como .exe empacotado pelo PyInstaller, precisa configurar
o caminho do Playwright Chromium antes de importar o servidor.
"""
import os
import sys
import webbrowser
import threading
import time


def setup_playwright_path():
    """Configura o caminho do Chromium quando rodando como .exe."""
    if getattr(sys, "frozen", False):
        # Rodando como .exe empacotado
        base_path = sys._MEIPASS
        playwright_path = os.path.join(base_path, "ms-playwright")
        if os.path.isdir(playwright_path):
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = playwright_path
            print(f"[LAUNCHER] Playwright browsers em: {playwright_path}")


def abrir_navegador_depois(url, delay=3):
    """Abre o navegador padrão após X segundos."""
    time.sleep(delay)
    try:
        webbrowser.open(url)
    except:
        pass


def main():
    setup_playwright_path()

    print("=" * 60)
    print("  AUTOMACAO DE BAIXA DE NF - DEALER.NET")
    print("=" * 60)
    print()
    print("  Servidor rodando em: http://localhost:5000")
    print("  Abrindo navegador automaticamente...")
    print()
    print("  NAO FECHE ESTA JANELA enquanto usar a ferramenta!")
    print("  Para encerrar, feche esta janela ou pressione CTRL+C")
    print("=" * 60)
    print()

    # Abre o navegador em background
    threading.Thread(
        target=abrir_navegador_depois,
        args=("http://localhost:5000",),
        daemon=True
    ).start()

    # Importa e roda o servidor (precisa ser DEPOIS do setup_playwright_path)
    from servidor import app
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
