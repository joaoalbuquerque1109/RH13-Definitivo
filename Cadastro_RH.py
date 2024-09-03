import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime

# Funções para manipulação do banco de dados
def connect_db():
    return sqlite3.connect('pessoas.db')

def create_table():
    with connect_db() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS pessoas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT,
                idade TEXT,
                data_nascimento TEXT,
                rg TEXT,
                cpf TEXT,
                endereco TEXT,
                bairro TEXT,
                cidade TEXT,
                estado TEXT,
                indicacao TEXT,
                telefone TEXT,
                email TEXT,
                estado_civil TEXT,
                pis_pasep TEXT,
                titulo_eleitor TEXT,
                zona TEXT,
                sessao TEXT,
                funcao TEXT,
                setor TEXT,
                data_admissao TEXT,
                salario TEXT,
                tipo TEXT,
                expediente1 TEXT
            )
        ''')
        conn.commit()

def add_person():
    data = (
        entry_nome.get() or None,
        entry_idade.get() or None,
        entry_data_nascimento.get() or None,
        entry_rg.get() or None,
        entry_cpf.get() or None,
        entry_endereco.get() or None,
        entry_bairro.get() or None,
        entry_cidade.get() or None,
        entry_estado.get() or None,
        entry_indicacao.get() or None,
        entry_telefone.get() or None,
        entry_email.get() or None,
        entry_estado_civil.get() or None,
        entry_pis_pasep.get() or None,
        entry_titulo_eleitor.get() or None,
        entry_zona.get() or None,
        entry_sessao.get() or None,
        entry_funcao.get() or None,
        entry_setor.get() or None,
        entry_data_admissao.get() or None,
        entry_salario.get() or None,
        entry_tipo.get() or None,
        entry_expediente1.get() or None
    )

    if not data[0]:
        messagebox.showwarning("Aviso", "Os campos 'Nome' é obrigatórios.")
        return
#REGRA DE 11 DÍGITOS NO CPF
    cpf = data[4]
    if len(cpf) != 11:
        messagebox.showwarning("Aviso", "O CPF deve ter exatamente 11 dígitos numéricos.")
        return

    with connect_db() as conn:
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO pessoas (
                    nome, idade, data_nascimento, rg, cpf, endereco, bairro, cidade, estado, indicacao, telefone,
                    email, estado_civil, pis_pasep, titulo_eleitor, zona, sessao, funcao, setor, data_admissao, 
                    salario, tipo, expediente1
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', data)
            conn.commit()
            messagebox.showinfo("Info", "Pessoa adicionada com sucesso.")
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao adicionar pessoa: {e}")

    clear_entries()
    load_people()

def clear_entries():
    for entry in entries.values():
        entry.delete(0, tk.END)

def load_people():
    with connect_db() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM pessoas')
        rows = cursor.fetchall()

    listbox_people.delete(*listbox_people.get_children())

    for row in rows:
        listbox_people.insert('', tk.END, iid=row[0], values=row[1:])

def delete_person():
    selected_item = listbox_people.selection()

    if not selected_item:
        messagebox.showwarning("Aviso", "Selecione um item para excluir.")
        return

    person_id = int(selected_item[0])

    with connect_db() as conn:
        cursor = conn.cursor()
        try:
            cursor.execute('DELETE FROM pessoas WHERE id = ?', (person_id,))
            conn.commit()
            messagebox.showinfo("Info", "Pessoa excluída com sucesso.")
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao excluir pessoa: {e}")

    load_people()

def export_to_excel():
    with connect_db() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM pessoas')
        rows = cursor.fetchall()

    if not rows:
        messagebox.showwarning("Aviso", "Nenhum dado para exportar.")
        return

    df = pd.DataFrame(rows, columns=[desc[0] for desc in cursor.description])

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Info", "Dados exportados para Excel com sucesso.")

def update_person():
    selected_item = listbox_people.selection()

    if not selected_item:
        messagebox.showwarning("Aviso", "Selecione um item para alterar.")
        return

    person_id = int(selected_item[0])

    data = (
        entry_nome.get() or None,
        entry_idade.get() or None,
        entry_data_nascimento.get() or None,
        entry_rg.get() or None,
        entry_cpf.get() or None,
        entry_endereco.get() or None,
        entry_bairro.get() or None,
        entry_cidade.get() or None,
        entry_estado.get() or None,
        entry_indicacao.get() or None,
        entry_telefone.get() or None,
        entry_email.get() or None,
        entry_estado_civil.get() or None,
        entry_pis_pasep.get() or None,
        entry_titulo_eleitor.get() or None,
        entry_zona.get() or None,
        entry_sessao.get() or None,
        entry_funcao.get() or None,
        entry_setor.get() or None,
        entry_data_admissao.get() or None,
        entry_salario.get() or None,
        entry_tipo.get() or None,
        entry_expediente1.get() or None,
        person_id
    )

    with connect_db() as conn:
        cursor = conn.cursor()
        try:
            cursor.execute('''
                UPDATE pessoas
                SET nome = ?, idade = ?, data_nascimento = ?, rg = ?, cpf = ?, endereco = ?, bairro = ?, cidade = ?, estado = ?, indicacao = ?, telefone = ?, email = ?, estado_civil = ?, pis_pasep = ?, titulo_eleitor = ?, zona = ?, sessao = ?, funcao = ?, setor = ?, data_admissao = ?, salario = ?, tipo = ?, expediente1 = ?
                WHERE id = ?
            ''', data)
            conn.commit()
            messagebox.showinfo("Info", "Pessoa alterada com sucesso.")
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao alterar pessoa: {e}")

    clear_entries()
    load_people()
#TESTE DE SELECIONAR
def on_select_item(event):
    selected_item = listbox_people.selection()

    if selected_item:
        item_id = int(selected_item[0])

        with connect_db() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM pessoas WHERE id = ?', (item_id,))
            person = cursor.fetchone()

        if person:
            for i, label in enumerate(labels):
                entry = entries[label]
                entry.delete(0, tk.END)
                entry.insert(0, person[i + 1])
#FIM DO TESTE
def search_person():
    search_term = entry_search.get().strip()
    if not search_term:
        load_people()
        messagebox.showwarning("Aviso", "Digite um termo de pesquisa.")
        return

    with connect_db() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT * FROM pessoas
            WHERE telefone LIKE ? OR cpf LIKE ?
        ''', (f'%{search_term}%', f'%{search_term}%'))
        rows = cursor.fetchall()

    listbox_people.delete(*listbox_people.get_children())

    for row in rows:
        listbox_people.insert('', tk.END, iid=row[0], values=row[1:])

def calculate_age(dob_str):
    try:
        dob = datetime.strptime(dob_str, '%d-%m-%Y')  # Ajuste o formato conforme necessário
        today = datetime.today()
        age = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
        return age
    except ValueError:
        return None

# Configuração da interface gráfica
root = tk.Tk()
root.title("Cadastro de Pessoas")
root.geometry("1200x800")
root.configure(bg="#F0F0F0")

# Estilo TTK
style = ttk.Style(root)
style.theme_use("clam")
style.configure("TLabel", font=("Segoe UI", 12), padding=6)
style.configure("TEntry", padding=6, font=("Segoe UI", 12), relief="flat", borderwidth=2) #, background="#F0F0F0"
style.configure("TButton", padding=6, relief="flat", background="#007ACC", foreground="white", font=("Segoe UI", 10))
style.map("TButton", background=[('active', '#005BA1')])
style.configure("Treeview", font=("Segoe UI", 12), background="#FFFFFF", fieldbackground="#FFFFFF", rowheight=25)
style.configure("Treeview.Heading", font=("Segoe UI", 13, "bold"))



# Frame para o formulário de cadastro
frame_form = ttk.Frame(root)
frame_form.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))


# Frame para pesquisa
frame_search = ttk.Frame(root)
frame_search.grid(row=3, column=0,pady=20, sticky=(tk.W, tk.E, tk.N, tk.S))

# Frame para a lista de pessoas
frame_list = ttk.Frame(root)
frame_list.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Frame para os botões
frame_buttons = ttk.Frame(root)
frame_buttons.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=2)
root.rowconfigure(2, weight=1)

# Definição das Labels e Entradas
labels = ["Nome", "Idade", "Data de Nascimento", "RG", "CPF", "Endereço", "Bairro", "Cidade", "Estado", "Indicação",
          "Telefone", "Email", "Estado Civil", "PIS/PASEP", "Título de Eleitor", "Zona", "Sessão", "Função", "Setor",
          "Data de Admissão", "Salário", "Tipo", "Expediente"]

entries = {}

for i, label in enumerate(labels):
    ttk.Label(frame_form, text=label).grid(row=i//4, column=i%4*2, padx=10, pady=8, sticky=tk.W)
    entry = ttk.Entry(frame_form, width=30)
    entry.grid(row=i//4, column=i%4*2+1, padx=10, pady=8, sticky=tk.W)
    entries[label] = entry

# Criação dos campos de entrada
entry_nome = entries["Nome"]
entry_idade = entries["Idade"]
entry_data_nascimento = entries["Data de Nascimento"]
entry_rg = entries["RG"]
entry_cpf = entries["CPF"]
entry_endereco = entries["Endereço"]
entry_bairro = entries["Bairro"]
entry_cidade = entries["Cidade"]
entry_estado = entries["Estado"]
entry_indicacao = entries["Indicação"]
entry_telefone = entries["Telefone"]
entry_email = entries["Email"]
entry_estado_civil = entries["Estado Civil"]
entry_pis_pasep = entries["PIS/PASEP"]
entry_titulo_eleitor = entries["Título de Eleitor"]
entry_zona = entries["Zona"]
entry_sessao = entries["Sessão"]
entry_funcao = entries["Função"]
entry_setor = entries["Setor"]
entry_data_admissao = entries["Data de Admissão"]
entry_salario = entries["Salário"]
entry_tipo = entries["Tipo"]
entry_expediente1 = entries["Expediente"]

# Campo de pesquisa
ttk.Label(frame_search, text="Pesquisar por Telefone ou CPF").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
entry_search = ttk.Entry(frame_search, width=30)
entry_search.grid(row=0, column=1, padx=10, pady=10, sticky=tk.W)
ttk.Button(frame_search, text="Pesquisar", command=search_person).grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)

# Botões de Ação
ttk.Button(frame_buttons, text="Adicionar", command=add_person).grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
ttk.Button(frame_buttons, text="Alterar", command=update_person).grid(row=0, column=1, padx=10, pady=10, sticky=tk.W)
ttk.Button(frame_buttons, text="Excluir", command=delete_person).grid(row=0, column=2, padx=10, pady=10, sticky=tk.W)
ttk.Button(frame_buttons, text="Exportar para Excel", command=export_to_excel).grid(row=0, column=3, padx=10, pady=10, sticky=tk.W)

# Lista de Pessoas
listbox_people = ttk.Treeview(frame_list, columns=labels, show='headings', selectmode="browse")
listbox_people.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

for label in labels:
    listbox_people.heading(label, text=label)
    listbox_people.column(label, width=120, anchor=tk.W)

listbox_people.bind("<<TreeviewSelect>>", on_select_item)

# Adicionar barra de rolagem
vsb = ttk.Scrollbar(frame_list, orient="vertical", command=listbox_people.yview)
vsb.grid(row=0, column=1, sticky='ns')
listbox_people.configure(yscrollcommand=vsb.set)

hsb = ttk.Scrollbar(frame_list, orient="horizontal", command=listbox_people.xview)
hsb.grid(row=1, column=0, sticky='ew')
listbox_people.configure(xscrollcommand=hsb.set)

frame_list.columnconfigure(0, weight=1)
frame_list.rowconfigure(0, weight=1)

# Inicializar a tabela e carregar os dados
create_table()
load_people()

root.mainloop()
