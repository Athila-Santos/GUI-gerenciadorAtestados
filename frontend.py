import sqlite3, os, sys, shutil, re, tkinter as tk
from tkinter import ttk, messagebox, filedialog
from unidecode import unidecode
from PyPDF2 import PdfFileReader
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from back.atestados_manager import AtestadosManager
import front.insert

class Frontend(tk.Tk):
    def __init__(self):
        super().__init__()
        default = os.path.join(os.getcwd(),"front", "start.ico")
        self.iconbitmap(default=default)
        self.title("Banco de Atestados de Capacidade Técnica")
        self.configure(background="#004080")
        self.geometry("1100x650")
        self.resizable(True, False)
        self.backend = AtestadosManager()
        
        self.create_widgets()
        self.open_treeview()  # Exibir todos os dados quando o widget for aberto

    def create_widgets(self):
        self.create_frames()
        self.create_buttons()
        self.create_labels()
        self.create_comboboxes()
        self.create_logo()
        self.create_list_frame()

    def create_frames(self):
        self.frame_1 = tk.Frame(self, bd=4, bg='#dfe3ee')
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.4)
        self.frame_2 = tk.Frame(self, bd=4, bg='#dfe3ee')
        self.frame_2.place(relx=0.02, rely=0.45, relwidth=0.96, relheight=0.5)

    def create_buttons(self):
        self.bt_search = tk.Button(self.frame_1, text="Buscar", font=("Helvetica", 13, "italic"),
                                   fg="white", bg="#FF8C00", command=self.search_data)
        self.bt_search.place(relx=0.875, rely=0.5, relwidth=0.12, relheight=0.09)

        self.bt_clear = tk.Button(self.frame_1, text="Limpar", font=("Helvetica", 13, "italic"),
                                   fg="white", bg="#FF8C00", command=self.clear_screen)
        self.bt_clear.place(relx=0.875, rely=0.6, relwidth=0.12, relheight=0.09)

        self.bt_insert = tk.Button(self.frame_1, text="Inserir", font=("Helvetica", 13, "italic"),
                                   fg="white", bg="#FF8C00", command=self.open_insert_window)
        self.bt_insert.place(relx=0.875, rely=0.7, relwidth=0.12, relheight=0.09)

        self.bt_export = tk.Button(self.frame_1, text="Ver na pasta", font=("Helvetica", 13, "italic"),
                                   fg="white", bg="#FF8C00", command=self.export_atestados)
        self.bt_export.place(relx=0.875, rely=0.8, relwidth=0.12, relheight=0.09)

    def create_labels(self):
        labels_data = [("Ano", 0.5), ("Emissor", 0.6), ("Empresa", 0.7), ("Participante", 0.8), ("Palavra-chave", 0.9)]
        for label_text, y_pos in labels_data:
            label = tk.Label(self.frame_1, text=label_text, font=("Helvetica", 12, "bold"), bg="#dfe3ee")
            label.place(relx=0.01, rely=y_pos)

        title_label = tk.Label(self.frame_1, text="GERENCIADOR DE ATESTADOS", font=("Helvetica", 20, "bold"), bg="#dfe3ee", fg="#004080")
        title_label.place(relx=0.5, rely=0.02, anchor="n")

    def create_comboboxes(self):
        # Dropdown para o Ano
        anos = self.backend.get_anos()
        self.selection_year = ttk.Combobox(self.frame_1, values=anos)
        self.selection_year.place(relx=0.13, rely=0.5, relwidth=0.3, relheight=0.075)

        # Dropdown para o Emissor
        emissores = self.backend.get_emissores()
        self.selection_client = ttk.Combobox(self.frame_1, values=emissores)
        self.selection_client.place(relx=0.13, rely=0.6, relwidth=0.3, relheight=0.075)

        # Entrada para a Empresa
        empresas = self.backend.get_empresas()
        self.selection_description = ttk.Combobox(self.frame_1, values=empresas)
        self.selection_description.place(relx=0.13, rely=0.7, relwidth=0.3, relheight=0.075)

        # Entrada para o Participante
        participantes = self.backend.get_participantes()
        self.participant_space = ttk.Combobox(self.frame_1, values=participantes)
        self.participant_space.place(relx=0.13, rely=0.8, relwidth=0.3, relheight=0.075)

        # Entrada para a Palavra-chave
        self.keyword = tk.Entry(self.frame_1, relief='groove')
        self.keyword.place(relx=0.13, rely=0.9, relwidth=0.3, relheight=0.075)

    def create_logo(self):
        file = os.path.join(os.getcwd(),"front", "imagem_ceplan.png")
        self.image = tk.PhotoImage(file=file)
        label_logo = tk.Label(self.frame_1, image=self.image, bg="#dfe3ee")
        label_logo.place(x=3, y=2)

    def create_list_frame(self):
        self.client_list = ttk.Treeview(self.frame_2, height=3, columns=["Coluna 1", "Coluna 2",
                                                                         "Coluna 3", "Coluna 4",
                                                                         "Coluna 5", "Coluna 6",
                                                                         "Coluna 7"])

        columns = ["ID","Código", "Ano", "Emissor", "Cliente", "Serviço Prestado", "Participante"]
        for i, col in enumerate(columns, start=1):
            self.client_list.heading(f"#{i}", text=col)

        self.client_list.column("#0", width=0, stretch=False)
        self.client_list.column("#1", width=30, stretch=False)  # Coluna 'id' com largura para três números
        self.client_list.column("#2", width=160, stretch=False)
        self.client_list.column("#3", width=40, stretch=False)
        self.client_list.column("#4", width=70, stretch=False)
        self.client_list.column("#5", width=30, stretch=True)
        self.client_list.column("#6", width=150, stretch=True)
        self.client_list.column("#7", width=150, stretch=True)


        self.client_list.place(relx=0, rely=0, relwidth=0.97, relheight=1)

        scroll_list = tk.Scrollbar(self.frame_2, orient="vertical", command=self.client_list.yview)
        scroll_list.place(relx=0.97, rely=0, relwidth=0.03, relheight=1)
        self.client_list.configure(yscrollcommand=scroll_list.set)

        style = ttk.Style(self)
        style.theme_use('alt')
        style.configure("Treeview.Heading", font=("Helvetica", 12, "bold"))

        # Associa o evento de seleção na Treeview ao método on_treeview_select
        self.client_list.bind('<Double-1>', self.on_treeview_select)

    def clear_screen(self):
        self.selection_year.set("")  # Limpa a seleção do ano
        self.selection_client.set("")  # Limpa a seleção do cliente
        self.selection_description.set("")  # Limpa a seleção da empresa
        self.participant_space.set("")  # Limpa a seleção do participante
        #self.keyword.delete(0, tk.END)
        self.client_list.delete(*self.client_list.get_children())
        self.open_treeview()  # Atualiza o Treeview com os dados originais

    def search_data(self):
        filtered_data = self.backend.get_filtered_data(self.selection_year, self.selection_client,
                                                       self.selection_description, self.participant_space, self.keyword)
        self.update_treeview(filtered_data)

    def export_atestados(self):
        filtered_data = self.backend.get_filtered_data(self.selection_year, self.selection_client,
                                                       self.selection_description, self.participant_space, self.keyword)
        self.backend.exportar_atestados(filtered_data)

    def open_insert_window(self):
        db_path = os.path.join(os.getcwd(), "back", "atestados.db")
        banco = AtestadosManager(db_path)
        app = front.insert.InterfaceTk(banco)
        app.mainloop()

    def open_treeview(self):
        all_data = self.backend.filtro_multiplo()
        # Popula a Treeview com os dados do banco de dados
        for row in all_data:
            self.client_list.insert('', 'end', values=row)
    
    def update_treeview(self, data=None):
        # Limpa a Treeview
        self.client_list.delete(*self.client_list.get_children())

        if data:
            for row in data:
                self.client_list.insert('', 'end', values=row)
        else:
            messagebox.showinfo("Atenção!", "Nenhum resultado para o filtro de busca selecionado!")
            self.clear_screen()

    def on_treeview_select(self, event):
        # Chama o método abrir_pdf da instância de Backend
        self.backend.abrir_pdf(event)
