import tkinter as tk, os, shutil
from tkinter import ttk, messagebox, filedialog

def limpar_texto(texto):
    return texto.upper().strip()

class InterfaceTk(tk.Tk):
    def __init__(self, banco):
        super().__init__()
        self.banco = banco
        default = os.path.join(os.getcwd(), "front", "start.ico")
        self.iconbitmap(default=default)
        self.title("Inserir Atestado")
        self.configure(background="#dfe3ee")
        self.criar_componentes()
        self.resizable(False, False)
        self.geometry("550x350")  # Redimensiona a janela principal
        self.participantes_atestado = []

    def criar_componentes(self):
        self.label_dia = tk.Label(self, text="Dia", font=("Helvetica", 9, "bold"), bg="#dfe3ee")
        self.label_dia.place(x=10, y=10)
        self.entry_dia = tk.Entry(self)
        self.entry_dia.place(x=120, y=10)

        self.label_mes = tk.Label(self, text="Mês", font=("Helvetica", 9, "bold"), bg="#dfe3ee")
        self.label_mes.place(x=10, y=40)
        self.entry_mes = tk.Entry(self)
        self.entry_mes.place(x=120, y=40)

        self.label_ano = tk.Label(self, text="Ano", font=("Helvetica", 9, "bold"), bg="#dfe3ee")
        self.label_ano.place(x=10, y=70)
        self.entry_ano = tk.Entry(self)
        self.entry_ano.place(x=120, y=70)

        self.label_emissor = tk.Label(self, text="Emissor", font=("Helvetica", 9, "bold"), bg="#dfe3ee")
        self.label_emissor.place(x=10, y=100)
        self.entry_emissor = tk.Entry(self)
        self.entry_emissor.place(x=120, y=100)

        self.label_cliente = tk.Label(self, text="Cliente", font=("Helvetica", 9, "bold"), bg="#dfe3ee")
        self.label_cliente.place(x=10, y=130)
        self.entry_cliente = tk.Entry(self)
        self.entry_cliente.place(x=120, y=130)

        self.label_servico_prestado = tk.Label(self, text="Serviço Prestado", font=("Helvetica", 9, "bold"), bg="#dfe3ee")
        self.label_servico_prestado.place(x=10, y=160)
        self.entry_servico_prestado = tk.Entry(self)
        self.entry_servico_prestado.place(x=120, y=160)

        self.label_participantes = tk.Label(self, text="Participantes", font=("Helvetica", 9, "bold"), bg="#dfe3ee")
        self.label_participantes.place(x=10, y=190)

        participantes_list = self.banco.get_participantes()

        self.entry_participantes = ttk.Combobox(self, values=participantes_list)
        self.entry_participantes.place(x=120, y=190, relwidth=0.55, relheight=0.05)

        self.button_adicionar_participante = tk.Button(self, text="Adicionar Participante", font=("Helvetica", 10, "italic"), fg="white", bg="#FF8C00", command=self.adicionar_participante)
        self.button_adicionar_participante.place(x=350, y=40, relheight=0.05, relwidth=0.3)

        self.label_participantes_adicionados = tk.Label(self, text="Participantes Adicionados", bg='#dfe3ee')
        self.label_participantes_adicionados.place(x=10, y=220)

        self.participantes_adicionados_text = tk.Text(self, height=5, width=40)
        self.participantes_adicionados_text.place(x=10, y=250)

        self.button_remover_participante = tk.Button(self, text="Remover Participantes", font=("Helvetica", 10, "italic"), fg="white", bg="#FF8C00", command=self.remover_participante)
        self.button_remover_participante.place(x=350, y=70, relheight=0.05, relwidth=0.3)

        self.button_adicionar_pdf = tk.Button(self, text="Anexar PDF", font=("Helvetica", 10, "italic"), fg="white", bg="#FF8C00", command=self.anexar_pdf)
        self.button_adicionar_pdf.place(x=350, y=100, relheight=0.05, relwidth=0.3)

        self.button_adicionar_atestado = tk.Button(self, text="Adicionar Atestado", font=("Helvetica", 10, "italic"), fg="white", bg="#FF8C00", command=self.adicionar_atestado)
        self.button_adicionar_atestado.place(x=350, y=130, relheight=0.05, relwidth=0.3)

        # Novos componentes no canto inferior direito
        self.button_alterar_registro = tk.Button(self, text="Alterar Registro", font=("Helvetica", 10, "italic"), fg="white", bg="#FF8C00", command=self.alterar_registro)
        self.button_alterar_registro.place(x=350, y=230, relheight=0.05, relwidth=0.3)

        self.button_fechar = tk.Button(self, text="Fechar", font=("Helvetica", 10, "italic"), fg="white", bg="#FF8C00", command=self.fechar_janela)
        self.button_fechar.place(x=350, y=260, relheight=0.05, relwidth=0.3)

    def adicionar_participante(self):
        participante = self.entry_participantes.get().strip()
        if participante:
            participante = limpar_texto(participante)
            try:
                if not self.banco.participante_existe(participante):
                    self.banco.adicionar_participante(participante)
                    messagebox.showinfo("Sucesso", f"Novo participante '{participante}' adicionado com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao adicionar participante: {e}")
            self.atualizar_lista_participantes_adicionados(participante)
            self.entry_participantes.delete(0, tk.END)

    def atualizar_lista_participantes_adicionados(self, participante):
        self.participantes_adicionados_text.insert(tk.END, participante.upper() + ", ")

    def remover_participante(self):
        self.participantes_adicionados_text.delete('1.0', 'end-2c')

    def anexar_pdf(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        dia = self.entry_dia.get().strip()
        mes = self.entry_mes.get().strip()
        ano = self.entry_ano.get().strip()
        if filepath:
            data = f"{ano}{mes.zfill(2)}{dia.zfill(2)}"
            emissor = limpar_texto(self.entry_emissor.get().strip())
            codigo = self.banco.obter_proximo_codigo(data, emissor)[-7:]

            if emissor in ["CLIENTE", "CORECON"]:
                pasta = emissor
                os.makedirs(pasta, exist_ok=True)
                shutil.copy(filepath, os.path.join(pasta, f"{codigo}.pdf"))
                messagebox.showinfo("Sucesso", "PDF anexado e salvo com sucesso!")
            else:
                messagebox.showerror("Erro", "Tipo de emissor inválido.")

    def adicionar_atestado(self):
        dia = self.entry_dia.get().strip()
        mes = self.entry_mes.get().strip()
        ano = self.entry_ano.get().strip()
        emissor = limpar_texto(self.entry_emissor.get().strip())
        cliente = limpar_texto(self.entry_cliente.get().strip())
        servico_prestado = limpar_texto(self.entry_servico_prestado.get().strip())

        participantes_entry = self.participantes_adicionados_text.get('1.0', 'end').strip()

        if participantes_entry:
            participantes = [limpar_texto(p) for p in participantes_entry.split(',') if p.strip()]
        else:
            messagebox.showerror("Erro", "Adicione pelo menos um participante para adicionar um atestado.")
            return

        if not all([dia, mes, ano, emissor, cliente, servico_prestado, participantes]):
            messagebox.showerror("Erro", "Preencha todos os campos para adicionar um atestado.")
            return

        if not (1 <= int(mes) <= 12):
            messagebox.showerror("Erro", "Mês inserido é inválido.")
            return

        if not (1 <= int(dia) <= 31):
            messagebox.showerror("Erro", "Dia inserido é inválido.")
            return

        if ((mes in ["04", "06", "09", "11"]) and int(dia) > 30) or (mes == "2" and int(dia) > 29):
            messagebox.showerror("Erro", "Dia inserido é inválido para o mês escolhido.")
            return

        if emissor not in ["CLIENTE", "CORECON"]:
            messagebox.showerror("Erro", "Tipo de emissor inválido.")
            return

        data = f"{ano}{mes.zfill(2)}{dia.zfill(2)}"
        codigo = self.banco.obter_proximo_codigo(data, emissor)
        if codigo is None:
            return

        self.banco.adicionar_atestado(codigo, ano, emissor, cliente, servico_prestado, participantes)
        messagebox.showinfo("Sucesso", "Atestado adicionado com sucesso!")
        self.limpar_campos()

    def alterar_registro(self):
        codigos_list = self.banco.get_codigos_atestados()

        self.alterar_frame = tk.Toplevel(self)
        self.alterar_frame.title("Alterar Atestado")
        self.alterar_frame.geometry("550x150")
        self.alterar_frame.configure(background="#dfe3ee")

        self.label_codigo = tk.Label(self.alterar_frame, text="Código", font=("Helvetica", 9, "bold"), bg="#dfe3ee")
        self.label_codigo.place(x=10, y=10)
        self.combobox_codigo = ttk.Combobox(self.alterar_frame, values=codigos_list)
        self.combobox_codigo.place(x=120, y=10, width=250)

        self.button_buscar_codigo = tk.Button(self.alterar_frame, text="Buscar", font=("Helvetica", 10, "italic"), fg="white", bg="#FF8C00", command=lambda: self.buscar_codigo(self.combobox_codigo.get()))
        self.button_buscar_codigo.place(x=380, y=10, width=150)

    def buscar_codigo(self, codigo):
        registro = self.banco.obter_registro(codigo)
        if not registro:
            messagebox.showerror("Erro", "Registro não encontrado.")
            return

        self.entry_dia.delete(0, tk.END)
        self.entry_dia.insert(0, registro['dia'])
        self.entry_mes.delete(0, tk.END)
        self.entry_mes.insert(0, registro['mes'])
        self.entry_ano.delete(0, tk.END)
        self.entry_ano.insert(0, registro['ano'])
        self.entry_emissor.delete(0, tk.END)
        self.entry_emissor.insert(0, registro['emissor'])
        self.entry_cliente.delete(0, tk.END)
        self.entry_cliente.insert(0, registro['cliente'])
        self.entry_servico_prestado.delete(0, tk.END)
        self.entry_servico_prestado.insert(0, registro['servico_prestado'])

        self.participantes_adicionados_text.delete('1.0', tk.END)
        self.participantes_adicionados_text.insert('1.0', ', '.join(registro['participantes']))

        self.button_salvar_alteracao = tk.Button(self.alterar_frame, text="Salvar Alteração", font=("Helvetica", 10, "italic"), fg="white", bg="#FF8C00", command=lambda: self.salvar_alteracao(codigo))
        self.button_salvar_alteracao.place(x=380, y=50, width=150)

        self.button_excluir = tk.Button(self.alterar_frame, text="Excluir", font=("Helvetica", 10, "italic"), fg="white", bg="#FF8C00", command=lambda: self.excluir_registro(codigo))
        self.button_excluir.place(x=380, y=90, width=150)

    def salvar_alteracao(self, codigo):
        dia = self.entry_dia.get().strip()
        mes = self.entry_mes.get().strip()
        ano = self.entry_ano.get().strip()
        emissor = limpar_texto(self.entry_emissor.get().strip())
        cliente = limpar_texto(self.entry_cliente.get().strip())
        servico_prestado = limpar_texto(self.entry_servico_prestado.get().strip())

        participantes_entry = self.participantes_adicionados_text.get('1.0', 'end').strip()

        if participantes_entry:
            participantes = [limpar_texto(p) for p in participantes_entry.split('\n') if p.strip()]
        else:
            messagebox.showerror("Erro", "Adicione pelo menos um participante para alterar um atestado.")
            return

        if not all([dia, mes, ano, emissor, cliente, servico_prestado, participantes]):
            messagebox.showerror("Erro", "Preencha todos os campos para alterar um atestado.")
            return

        if not (1 <= int(mes) <= 12):
            messagebox.showerror("Erro", "Mês inserido é inválido.")
            return

        if not (1 <= int(dia) <= 31):
            messagebox.showerror("Erro", "Dia inserido é inválido.")
            return

        if ((mes in ["04", "06", "09", "11"]) and int(dia) > 30) or (mes == "02" and int(dia) > 29):
            messagebox.showerror("Erro", "Dia inserido é inválido para o mês escolhido.")
            return

        if emissor not in ["CLIENTE", "CORECON"]:
            messagebox.showerror("Erro", "Tipo de emissor inválido.")
            return

        data = f"{ano}{mes.zfill(2)}{dia.zfill(2)}"
        codigo_novo = self.banco.obter_proximo_codigo(data, emissor)
        if codigo_novo is None:
            return

        self.banco.atualizar_atestado(codigo, codigo_novo, ano, emissor, cliente, servico_prestado, participantes)
        messagebox.showinfo("Sucesso", "Atestado alterado com sucesso!")
        self.limpar_campos()

    def excluir_registro(self, codigo):
        if messagebox.askyesno("Confirmar", "Tem certeza que deseja excluir este registro?"):
            self.banco.excluir_atestado(codigo)
            messagebox.showinfo("Sucesso", "Atestado excluído com sucesso!")
            self.limpar_campos()
            self.alterar_frame.destroy()

    def fechar_janela(self):
        self.banco.fechar_conexao()
        self.destroy()

    def limpar_campos(self):
        self.entry_dia.delete(0, tk.END)
        self.entry_mes.delete(0, tk.END)
        self.entry_ano.delete(0, tk.END)
        self.entry_emissor.delete(0, tk.END)
        self.entry_cliente.delete(0, tk.END)
        self.entry_servico_prestado.delete(0, tk.END)
        self.participantes_adicionados_text.delete('1.0', tk.END)