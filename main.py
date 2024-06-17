from back.atestados_manager import AtestadosManager
from front.frontend import Frontend


if __name__ == "__main__":
    app = Frontend()
    app.mainloop()

# cmd: pyinstaller --onefile --noconsole --clean main.py
