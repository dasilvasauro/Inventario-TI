import customtkinter as ctk
import sqlite3
import os
from datetime import datetime
from tkinter import ttk, messagebox, filedialog
import tkinter as tk
from typing import List, Dict, Any, Optional
from PIL import Image, ImageTk

# Configuração do tema
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# Dados das comboboxes (chumbados no código)
COMBOBOX_WORKSTATIONS = {
    "Local": [
        "Welcome Desk", "Back Office WD", "Sala Gerente Recepção", "Sala Gerente Geral",
        "Guest Service", "CPD", "Sala Governança", "Sala Reservas", "Sala Gamer",
        "Sala P&C", "Sala Eventos", "Base Restaurante", "Sala Chef de Cozinha",
        "Sala A&B", "Sala Manutenção", "Sala Gerente Manutenção", "Sala Controladoria",
        "Sala Caixa Geral", "Sala Compras", "Sala Marketing", "Sala Assist. GG",
        "Sala Vendas", "Sala CX", "Almoxarifado", "Kollab", "Sala Gerente Operacional",
        "Sala Segurança", "Sala Banquetes", "Room Service", "Outro"
    ],
    "Sist. Oper.": [
        "Windows 10 Pro", "Windows 11 Pro"
    ],
    "Build": [
        "21H2", "22H2", "23H2", "24H2", "25H2", "26H2"
    ],
    "Bomgar":[
        "Sim","Não"
    ],
    "Modelo": [
        "Latitude 3400", "Latitude 3420", "Latitude 3440", "Latitude 3450", "Latitude 3470",
        "Latitude 3480", "Latitude 3540", "Optiplex 3000", "Optiplex 3060", "Optiplex 3070",
        "Optiplex 3090", "Optiplex 7010", "Optiplex 7010 SFF", "Optiplex 7020", "Optiplex 7070",
        "Optiplex 7450"
    ],
    "RAM": [
        "8GB", "16GB", "32GB", "64GB", "128GB"
    ],
    "Barramento": [
        "DDR3", "DDR4", "DDR5", "DDR6"
    ],
    "SSD": [
        "256GB", "512GB", "1TB"
    ],
    "Chipset": [
        "i5-9500", "i5-14500T", "i5-1135G7", "i5-10500T", "i5-8265U", "i5-13500T",
        "i3-9100T", "i5-12500T", "i3-8100T", "i5-6200U", "i5-1345U", "i5-13500",
        "Ultra 5 235A", "Ultra 7 265U"
    ]
}

COMBOBOX_ATIVOS = {
    "Tipo": [
        "Monitor", "Mouse", "Teclado", "Carregador", "HD Externo", "Cabo de Rede",
        "Telefone", "Celular", "Pendrive", "No-break", "Headset", "Webcam",
        "Mi Stick", "Access Point", "Hub USB", "Videogame"
    ],
    "Marca": [
        "Dell", "Samsung", "Logitech", "LG", "AOC", "HP", "Microsoft", "Intelbras",
        "Lenovo", "Baseus", "Soho Plus", "Schneider", "SMS", "APC", "EPSON"
    ],
    "Estado": [
        "Novo", "Em uso", "Usado", "Avariado", "P/ Descarte"
    ]
}

class Database:
    def __init__(self, db_path):
        self.db_path = db_path
        self.init_db()
    
    def get_connection(self):
        return sqlite3.connect(self.db_path)
    
    def init_db(self):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Tabela de workstations
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS workstations (
            hostname TEXT PRIMARY KEY,
            usuario TEXT,
            local TEXT NOT NULL,
            sist_oper TEXT NOT NULL,
            build TEXT NOT NULL,
            bomgar TEXT NOT NULL,
            service_tag TEXT NOT NULL,
            mac TEXT NOT NULL,
            modelo TEXT NOT NULL,
            ram TEXT NOT NULL,
            barramento TEXT,
            ssd TEXT NOT NULL,
            chipset TEXT NOT NULL,
            data_compra TEXT NOT NULL
        )
        ''')
        
        # Tabela de ativos
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS ativos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tipo TEXT NOT NULL,
            marca TEXT NOT NULL,
            modelo TEXT NOT NULL,
            estado TEXT NOT NULL,
            notas TEXT,
            data_compra TEXT,
            maquina_associada TEXT,
            FOREIGN KEY (maquina_associada) REFERENCES workstations (hostname)
        )
        ''')
        
        conn.commit()
        conn.close()
    
    def get_workstations(self, filters=None):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        query = "SELECT * FROM workstations"
        params = []
        
        if filters and filters.get('field') and filters.get('value'):
            query += f" WHERE {filters['field']} LIKE ?"
            params.append(f"%{filters['value']}%")
        
        query += " ORDER BY hostname"

        cursor.execute(query, params)
        columns = [description[0] for description in cursor.description]
        results = cursor.fetchall()
        
        conn.close()
        
        # Calcular tempo de uso para cada workstation
        workstations = []
        for row in results:
            workstation = dict(zip(columns, row))
            workstation['tempo_uso'] = self.calcular_tempo_uso(workstation['data_compra'])
            workstations.append(workstation)
        
        return workstations
    
    def get_ativos(self, filters=None, maquina_associada=None):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        query = "SELECT * FROM ativos"
        params = []
        
        conditions = []
        if filters and filters.get('field') and filters.get('value'):
            conditions.append(f"{filters['field']} LIKE ?")
            params.append(f"%{filters['value']}%")
        
        if maquina_associada:
            conditions.append("maquina_associada = ?")
            params.append(maquina_associada)
        
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        
        cursor.execute(query, params)
        columns = [description[0] for description in cursor.description]
        results = cursor.fetchall()
        
        conn.close()
        
        # Calcular tempo de uso para cada ativo
        ativos = []
        for row in results:
            ativo = dict(zip(columns, row))
            if ativo['data_compra']:
                ativo['tempo_uso'] = self.calcular_tempo_uso(ativo['data_compra'])
            else:
                ativo['tempo_uso'] = "N/A"
            ativos.append(ativo)
        
        return ativos
    
    def insert_workstation(self, data):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
            INSERT INTO workstations 
            (hostname, usuario, local, sist_oper, build, bomgar, service_tag, mac, modelo, ram, barramento, ssd, chipset, data_compra)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data['hostname'], data['usuario'], data['local'], data['sist_oper'], 
                data['build'], data['bomgar'], data['service_tag'], data['mac'], data['modelo'], 
                data['ram'], data['barramento'], data['ssd'], data['chipset'], data['data_compra']
            ))
            
            conn.commit()
            conn.close()
            return True
        except sqlite3.IntegrityError:
            conn.close()
            return False
    
    def insert_ativo(self, data):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
            INSERT INTO ativos 
            (tipo, marca, modelo, estado, notas, data_compra, maquina_associada)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                data['tipo'], data['marca'], data['modelo'], data['estado'], 
                data['notas'], data['data_compra'], data['maquina_associada']
            ))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            conn.close()
            return False
    
    def update_workstation(self, hostname, data):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
            UPDATE workstations 
            SET usuario=?, local=?, sist_oper=?, build=?, bomgar=?, service_tag=?, mac=?, modelo=?, ram=?, barramento=?, ssd=?, chipset=?, data_compra=?
            WHERE hostname=?
            ''', (
                data['usuario'], data['local'], data['sist_oper'], data['build'], data['bomgar'], 
                data['service_tag'], data['mac'], data['modelo'], data['ram'], 
                data['barramento'], data['ssd'], data['chipset'], data['data_compra'], hostname
            ))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            conn.close()
            return False
    
    def update_ativo(self, id, data):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
            UPDATE ativos 
            SET tipo=?, marca=?, modelo=?, estado=?, notas=?, data_compra=?, maquina_associada=?
            WHERE id=?
            ''', (
                data['tipo'], data['marca'], data['modelo'], data['estado'], 
                data['notas'], data['data_compra'], data['maquina_associada'], id
            ))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            conn.close()
            return False
    
    def delete_workstation(self, hostname):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            # Primeiro, desassociar quaisquer ativos vinculados a esta workstation
            cursor.execute('UPDATE ativos SET maquina_associada = NULL WHERE maquina_associada = ?', (hostname,))
            
            # Depois, excluir a workstation
            cursor.execute('DELETE FROM workstations WHERE hostname = ?', (hostname,))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            conn.close()
            return False
    
    def delete_ativo(self, id):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            cursor.execute('DELETE FROM ativos WHERE id = ?', (id,))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            conn.close()
            return False
    
    def get_workstation(self, hostname):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM workstations WHERE hostname = ?', (hostname,))
        columns = [description[0] for description in cursor.description]
        result = cursor.fetchone()
        
        conn.close()
        
        if result:
            workstation = dict(zip(columns, result))
            workstation['tempo_uso'] = self.calcular_tempo_uso(workstation['data_compra'])
            return workstation
        return None
    
    def get_ativo(self, id):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM ativos WHERE id = ?', (id,))
        columns = [description[0] for description in cursor.description]
        result = cursor.fetchone()
        
        conn.close()
        
        if result:
            ativo = dict(zip(columns, result))
            if ativo['data_compra']:
                ativo['tempo_uso'] = self.calcular_tempo_uso(ativo['data_compra'])
            else:
                ativo['tempo_uso'] = "N/A"
            return ativo
        return None
    
    def calcular_tempo_uso(self, data_compra):
        try:
            compra = datetime.strptime(data_compra, "%d/%m/%Y")
            hoje = datetime.now()
            
            diff = hoje - compra
            anos = diff.days // 365
            meses = (diff.days % 365) // 30
            dias = (diff.days % 365) % 30
            
            return f"{anos} anos, {meses} meses e {dias} dias"
        except:
            return "Data inválida"

class DateEntry(ctk.CTkEntry):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.bind('<KeyRelease>', self.format_date)
    
    def format_date(self, event):
        text = self.get().replace("/", "")[:8]
        if len(text) > 4:
            self.delete(0, tk.END)
            self.insert(0, f"{text[:2]}/{text[2:4]}/{text[4:]}")
        elif len(text) > 2:
            self.delete(0, tk.END)
            self.insert(0, f"{text[:2]}/{text[2:]}")
        
        # Validar data
        self.validate_date()
    
    def validate_date(self):
        date_str = self.get()
        if len(date_str) == 10:
            try:
                day, month, year = map(int, date_str.split('/'))
                datetime(year, month, day)
                self.configure(border_color="#3B8ED0")  # Cor normal
            except ValueError:
                self.configure(border_color="red")  # Cor de erro
        else:
            self.configure(border_color="#3B8ED0")  # Cor normal

class AutocompleteCombobox(ctk.CTkComboBox):
    def __init__(self, master, values, **kwargs):
        super().__init__(master, values=values, **kwargs)
        self.values = values
        self.bind('<KeyRelease>', self.autocomplete)
    
    def autocomplete(self, event):
        current = self.get()
        if current:
            matches = [v for v in self.values if v.lower().startswith(current.lower())]
            if matches:
                self.configure(values=matches)
            else:
                self.configure(values=self.values)
        else:
            self.configure(values=self.values)

class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Inventário de TI - Pullman Guarulhos São Paulo Airport")
        self.geometry("1366x768")
        self.minsize(1366, 768)
        
        # Centralizar a janela
        self.center_window()
        self.state("zoomed")
        
        # Configuração do banco de dados
        self.db_path = None
        self.db = None
        
        # Carregar configurações
        self.load_config()
        
        # Se já temos um caminho de banco, inicializar
        if self.db_path and os.path.exists(self.db_path):
            self.db = Database(self.db_path)
            self.show_main_menu()
        else:
            self.show_db_selection()
    
    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def load_config(self):
        config_path = os.path.join(os.path.dirname(__file__), "config.ini")
        if os.path.exists(config_path):
            with open(config_path, 'r') as f:
                self.db_path = f.read().strip()
    
    def save_config(self):
        config_path = os.path.join(os.path.dirname(__file__), "config.ini")
        with open(config_path, 'w') as f:
            f.write(self.db_path)
    
    def show_db_selection(self):
        self.clear_window()
        
        frame = ctk.CTkFrame(self,fg_color="transparent")
        frame.pack(expand=True, fill="both", padx=50, pady=50)
        
        try:
            image_path = "newico.png"
            pil_image = Image.open(image_path)

            ctk_image = ctk.CTkImage(
                light_image=pil_image,
                dark_image=pil_image,
                size=(200, 200)
            )

            image_label = ctk.CTkLabel(frame, image=ctk_image, text="")
            image_label.pack(pady=(0, 20))
        except Exception as e:
            print(f"Erro ao carregar a imagem: {e}")
            error_label = ctk.CTkLabel(frame, text="📝", font=ctk.CTkFont(size=48))
            error_label.pack(pady=(0, 28))

        title_label = ctk.CTkLabel(frame, text="Inventário de Hardware - Pullman Guarulhos", 
                                 font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=28)
        subtitulo_db = ctk.CTkLabel(frame, text="Selecione ou crie um banco de dados para iniciar",
                                 font=ctk.CTkFont(size=16))
        subtitulo_db.pack(pady=15)
        
        # Botão para criar novo banco
        new_db_btn = ctk.CTkButton(frame, text="Criar Novo Banco de Dados", 
                                  command=self.create_new_db, height=50, fg_color="#98D3A5", hover_color="#79A883",border_color="#7FB18A", border_width=1, text_color="black")
        new_db_btn.pack(pady=20, padx=100, fill="x")
        
        # Botão para usar banco existente
        existing_db_btn = ctk.CTkButton(frame, text="Usar Banco de Dados Existente", 
                                       command=self.select_existing_db, height=50, fg_color="#F1F1F1", hover_color="#DADADA", border_color="#7C7C7C", border_width=1, text_color="black")
        existing_db_btn.pack(pady=20, padx=100, fill="x")
    
    def create_new_db(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".db",
            filetypes=[("SQLite Database", "*.db")],
            title="Criar Novo Banco de Dados"
        )
        
        if file_path:
            self.db_path = file_path
            self.db = Database(self.db_path)
            self.save_config()
            self.show_main_menu()
    
    def select_existing_db(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("SQLite Database", "*.db")],
            title="Selecionar Banco de Dados"
        )
        
        if file_path:
            self.db_path = file_path
            self.db = Database(self.db_path)
            self.save_config()
            self.show_main_menu()
    
    def show_main_menu(self):
        self.clear_window()
        self.state("zoomed")
        
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(expand=True, fill="both", padx=50, pady=50)
        
        try:
            image_path = "newico.png"
            pil_image = Image.open(image_path)

            ctk_image = ctk.CTkImage(
                light_image=pil_image,
                dark_image=pil_image,
                size=(200, 200)
            )

            image_label = ctk.CTkLabel(frame, image=ctk_image, text="")
            image_label.pack(pady=(0, 20))
        except Exception as e:
            print(f"Erro ao carregar a imagem: {e}")
            error_label = ctk.CTkLabel(frame, text="📝", font=ctk.CTkFont(size=48))
            error_label.pack(pady=(0, 28))

        title_label = ctk.CTkLabel(frame, text="Inventário de Hardware - Pullman Guarulhos", 
                                 font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=28)
        
        # Botão para gerenciar workstations
        workstations_btn = ctk.CTkButton(frame, text="Gestão de Workstations", 
                                        command=self.show_workstations,fg_color="#F1F1F1", hover_color="#DADADA", border_color="#7C7C7C", border_width=1, text_color="black", height=50)
        workstations_btn.pack(pady=20, padx=100, fill="x")
        
        # Botão para gerenciar ativos
        ativos_btn = ctk.CTkButton(frame, text="Outros Ativos", 
                                  command=self.show_ativos, fg_color="#F1F1F1", hover_color="#DADADA", border_color="#7C7C7C", border_width=1, text_color="black", height=50)
        ativos_btn.pack(pady=20, padx=100, fill="x")
        
        # Botão para voltar à seleção de banco
        back_btn = ctk.CTkButton(frame, text="Alterar Banco de Dados", 
                                command=self.show_db_selection, height=50, fg_color="#C5DAE4", hover_color="#99B6CA", border_color="#5B7E9E", border_width=1, text_color="black")
        back_btn.pack(pady=20, padx=100, fill="x")
    
    def show_workstations(self):
        WorkstationsWindow(self, self.db)
    
    def show_ativos(self):
        AtivosWindow(self, self.db)
    
    def clear_window(self):
        for widget in self.winfo_children():
            widget.destroy()

class BaseCRUDWindow(ctk.CTkToplevel):
    def __init__(self, master, db, title):
        super().__init__(master)
        
        self.db = db
        self.title(title)
        self.geometry("1366x768")
        self.minsize(1366, 768)
        self.center_window()
        self.state("zoomed")
        
        # Configurar protocolo de fechamento
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Frame principal
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Frame de pesquisa
        self.search_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.search_frame.pack(fill="x", padx=10, pady=10)
        
        # Frame da tabela
        self.table_frame = ctk.CTkFrame(self.main_frame)
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Frame de botões
        self.button_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.button_frame.pack(fill="x", padx=10, pady=10)
        
        # Botão voltar
        self.back_btn = ctk.CTkButton(self.button_frame, text="<  Voltar", 
                                     command=self.on_close, fg_color="#F1F1F1", hover_color="#DADADA", border_color="#7C7C7C", border_width=1, text_color="black")
        self.back_btn.pack(side="left", padx=5)

    def center_window(self):
        self.state("zoomed")
    
    def on_close(self):
        self.destroy()
        self.master.deiconify()

class WorkstationsWindow(BaseCRUDWindow):
    def __init__(self, master, db):
        super().__init__(master, db, "Gestão de Workstations")
        
        # Configurar interface
        self.setup_ui()
        
        # Carregar dados
        self.load_data()
        
        # Esconder janela principal
        self.master.withdraw()
    
    def setup_ui(self):
        # Campos de pesquisa
        ctk.CTkLabel(self.search_frame, text="Pesquisar:").pack(side="left", padx=5)
        
        self.search_field = ctk.CTkComboBox(self.search_frame, 
                                          values=["Hostname", "Usuario", "Local", "Sist. Oper.", 
                                                 "Build","Bomgar", "Service Tag", "MAC", "Modelo", "RAM", 
                                                 "Barramento", "SSD", "Chipset", "Data Compra"])
        self.search_field.pack(side="left", padx=5)
        self.search_field.set("Hostname")
        
        self.search_entry = ctk.CTkEntry(self.search_frame)
        self.search_entry.pack(side="left", padx=5, fill="x", expand=True)
        self.search_entry.bind("<KeyRelease>", self.on_search)
        
        # Botão atualizar
        self.update_btn = ctk.CTkButton(self.search_frame, text="Atualizar Lista", 
                                       command=self.load_data, fg_color="#F1F1F1", hover_color="#DADADA", border_color="#7C7C7C", border_width=1, text_color="black")
        self.update_btn.pack(side="right", padx=5)
        
        # Tabela
        self.setup_table()
        
        # Botões de ação
        self.add_btn = ctk.CTkButton(self.button_frame, text="+  Adicionar", 
                                    command=self.add_item, fg_color="#98D3A5", hover_color="#79A883",border_color="#7FB18A", border_width=1, text_color="black")
        self.add_btn.pack(side="left", padx=5)
        
        self.edit_btn = ctk.CTkButton(self.button_frame, text="Editar", 
                                     command=self.edit_item, fg_color="#F1F1F1", hover_color="#DADADA", border_color="#7C7C7C", border_width=1, text_color="black")
        self.edit_btn.pack(side="left", padx=5)
        
        self.duplicate_btn = ctk.CTkButton(self.button_frame, text="Duplicar", 
                                          command=self.duplicate_item, fg_color="#F1F1F1", hover_color="#DADADA", border_color="#7C7C7C", border_width=1, text_color="black")
        self.duplicate_btn.pack(side="left", padx=5)
        
        self.delete_btn = ctk.CTkButton(self.button_frame, text="-  Excluir", 
                                       command=self.delete_item, fg_color="#E6ABAB", hover_color="#DDA6A6", border_color="#805E5E", border_width=1, text_color="black")
        self.delete_btn.pack(side="left", padx=5)
    
    def setup_table(self):
        # Criar treeview com scrollbars
        self.tree_frame = ctk.CTkFrame(self.table_frame)
        self.tree_frame.pack(fill="both", expand=True)
        
        self.tree = ttk.Treeview(self.tree_frame, columns=(
            "hostname", "usuario", "local", "sist_oper", "build","bomgar", "service_tag", 
            "mac", "modelo", "ram", "barramento", "ssd", "chipset", "data_compra", "tempo_uso"
        ), show="headings", height=20, selectmode="extended")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)  

        # Definir colunas
        columns = {
            "hostname": "Hostname",
            "usuario": "Usuário",
            "local": "Local",
            "sist_oper": "Sist. Oper.",
            "build": "Build",
            "bomgar": "Bomgar",
            "service_tag": "Service Tag",
            "mac": "MAC",
            "modelo": "Modelo",
            "ram": "RAM",
            "barramento": "Barramento",
            "ssd": "SSD",
            "chipset": "Chipset",
            "data_compra": "Data Compra",
            "tempo_uso": "Tempo de Uso"
        }
        
        for col, text in columns.items():
            self.tree.heading(col, text=text)
            self.tree.column(col, width=100, minwidth=50)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        # Label para mostrar contagem de seleção
        self.selection_label = ctk.CTkLabel(self.tree_frame, text="0 item(ns) selecionado(s)")
        self.selection_label.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        # Vincular evento de seleção
        self.tree.bind('<<TreeviewSelect>>', self.update_selection_count)
        
        # Bind duplo clique
        self.tree.bind("<Double-1>", self.on_double_click)

    def update_selection_count(self, event=None):
        selected = len(self.tree.selection())
        self.selection_label.configure(text=f"{selected} item(ns) selecionado(s)")

    def load_data(self, filters=None):
        # Limpar tabela
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Carregar dados
        workstations = self.db.get_workstations(filters)
        
        for ws in workstations:
            self.tree.insert("", "end", values=(
                ws['hostname'], ws['usuario'], ws['local'], ws['sist_oper'], 
                ws['build'], ws['bomgar'], ws['service_tag'], ws['mac'], ws['modelo'], 
                ws['ram'], ws['barramento'], ws['ssd'], ws['chipset'], 
                ws['data_compra'], ws['tempo_uso']
            ))
    
    def on_search(self, event):
        field = self.search_field.get().lower().replace(" ", "_").replace(".", "")
        value = self.search_entry.get()
        
        if value:
            self.load_data({'field': field, 'value': value})
        else:
            self.load_data()
    
    def on_double_click(self, event):
        item = self.tree.selection()
        if item:
            hostname = self.tree.item(item[0])['values'][0]
            self.show_details(hostname)
    
    def show_details(self, hostname):
        DetailsWindow(self, self.db, hostname)
    
    def add_item(self):
        EditWorkstationWindow(self, self.db)
    
    def edit_item(self):
        item = self.tree.selection()
        if item:
            hostname = self.tree.item(item[0])['values'][0]
            EditWorkstationWindow(self, self.db, hostname)
        else:
            messagebox.showwarning("Seleção", "Selecione uma workstation para editar.")
    
    def duplicate_item(self):
        item = self.tree.selection()
        if item:
            hostname = self.tree.item(item[0])['values'][0]
            DuplicateWorkstationWindow(self, self.db, hostname)
        else:
            messagebox.showwarning("Seleção", "Selecione uma workstation para duplicar.")
    
    def delete_item(self):
        item = self.tree.selection()
        if item:
            hostname = self.tree.item(item[0])['values'][0]
            
            if messagebox.askyesno("Confirmar", f"Tem certeza que deseja excluir a workstation {hostname}?"):
                if self.db.delete_workstation(hostname):
                    messagebox.showinfo("Sucesso", "Workstation excluída com sucesso.")
                    self.load_data()
                else:
                    messagebox.showerror("Erro", "Não foi possível excluir a workstation.")
        else:
            messagebox.showwarning("Seleção", "Selecione uma workstation para excluir.")

class AtivosWindow(BaseCRUDWindow):
    def __init__(self, master, db):
        super().__init__(master, db, "Gestão de Outros Ativos")
        
        # Configurar interface
        self.setup_ui()
        
        # Carregar dados
        self.load_data()
        
        # Esconder janela principal
        self.master.withdraw()
    
    def setup_ui(self):
        # Campos de pesquisa
        ctk.CTkLabel(self.search_frame, text="Pesquisar:").pack(side="left", padx=5)
        
        self.search_field = ctk.CTkComboBox(self.search_frame, 
                                          values=["Tipo", "Marca", "Modelo", "Estado", 
                                                 "Notas", "Data Compra", "Maquina Associada"])
        self.search_field.pack(side="left", padx=5)
        self.search_field.set("Tipo")
        
        self.search_entry = ctk.CTkEntry(self.search_frame)
        self.search_entry.pack(side="left", padx=5, fill="x", expand=True)
        self.search_entry.bind("<KeyRelease>", self.on_search)
        
        # Botão atualizar
        self.update_btn = ctk.CTkButton(self.search_frame, text="Atualizar Lista", 
                                       command=self.load_data, fg_color="#F1F1F1", hover_color="#DADADA", border_color="#7C7C7C", border_width=1, text_color="black")
        self.update_btn.pack(side="right", padx=5)
        
        # Tabela
        self.setup_table()
        
        # Botões de ação
        self.add_btn = ctk.CTkButton(self.button_frame, text="+  Adicionar", 
                                    command=self.add_item, fg_color="#98D3A5", hover_color="#79A883",border_color="#7FB18A", border_width=1, text_color="black")
        self.add_btn.pack(side="left", padx=5)
        
        self.edit_btn = ctk.CTkButton(self.button_frame, text="Editar", 
                                     command=self.edit_item, fg_color="#F1F1F1", hover_color="#DADADA", border_color="#7C7C7C", border_width=1, text_color="black")
        self.edit_btn.pack(side="left", padx=5)
        
        self.delete_btn = ctk.CTkButton(self.button_frame, text="Excluir", 
                                       command=self.delete_item, fg_color="#E6ABAB", hover_color="#DDA6A6", border_color="#805E5E", border_width=1, text_color="black")
        self.delete_btn.pack(side="left", padx=5)
    
    def setup_table(self):
        # Criar treeview com scrollbars
        self.tree_frame = ctk.CTkFrame(self.table_frame)
        self.tree_frame.pack(fill="both", expand=True)

        
        self.tree = ttk.Treeview(self.tree_frame, columns=(
            "id", "tipo", "marca", "modelo", "estado", "notas", 
            "data_compra", "maquina_associada", "tempo_uso"
        ), show="headings", height=20, selectmode="extended")
        
        # Definir colunas
        columns = {
            "id": "ID",
            "tipo": "Tipo",
            "marca": "Marca",
            "modelo": "Modelo",
            "estado": "Estado",
            "notas": "Notas",
            "data_compra": "Data Compra",
            "maquina_associada": "Máquina Associada",
            "tempo_uso": "Tempo de Uso"
        }
        
        for col, text in columns.items():
            self.tree.heading(col, text=text)
            self.tree.column(col, width=100, minwidth=50)
        
        # Esconder coluna ID
        self.tree.column("id", width=0, stretch=False)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        # Label para mostrar contagem de seleção
        self.selection_label = ctk.CTkLabel(self.tree_frame, text="0 item(ns) selecionado(s)")
        self.selection_label.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        # Vincular evento de seleção
        self.tree.bind('<<TreeviewSelect>>', self.update_selection_count)
        
        # Bind duplo clique
        self.tree.bind("<Double-1>", self.on_double_click)

    def update_selection_count(self, event=None):
        selected = len(self.tree.selection())
        self.selection_label.configure(text=f"{selected} item(ns) selecionado(s)")
    
    def load_data(self, filters=None):
        # Limpar tabela
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Carregar dados
        ativos = self.db.get_ativos(filters)
        
        for ativo in ativos:
            self.tree.insert("", "end", values=(
                ativo['id'], ativo['tipo'], ativo['marca'], ativo['modelo'], 
                ativo['estado'], ativo['notas'], ativo['data_compra'], 
                ativo['maquina_associada'], ativo['tempo_uso']
            ))
    
    def on_search(self, event):
        field = self.search_field.get().lower().replace(" ", "_")
        value = self.search_entry.get()
        
        if value:
            self.load_data({'field': field, 'value': value})
        else:
            self.load_data()
    
    def on_double_click(self, event):
        item = self.tree.selection()
        if item:
            id = self.tree.item(item[0])['values'][0]
            self.show_details(id)
    
    def show_details(self, id):
        # Para ativos, podemos mostrar os detalhes em uma mensagem simples
        ativo = self.db.get_ativo(id)
        if ativo:
            details = f"Tipo: {ativo['tipo']}\nMarca: {ativo['marca']}\nModelo: {ativo['modelo']}\nEstado: {ativo['estado']}\nNotas: {ativo['notas']}\nData Compra: {ativo['data_compra']}\nMáquina Associada: {ativo['maquina_associada']}\nTempo de Uso: {ativo['tempo_uso']}"
            messagebox.showinfo("Detalhes do Ativo", details)
    
    def add_item(self):
        EditAtivoWindow(self, self.db)
    
    def edit_item(self):
        item = self.tree.selection()
        if item:
            id = self.tree.item(item[0])['values'][0]
            EditAtivoWindow(self, self.db, id)
        else:
            messagebox.showwarning("Seleção", "Selecione um ativo para editar.")
    
    def delete_item(self):
        item = self.tree.selection()
        if item:
            id = self.tree.item(item[0])['values'][0]
            tipo = self.tree.item(item[0])['values'][1]
            
            if messagebox.askyesno("Confirmar", f"Tem certeza que deseja excluir o ativo {tipo}?"):
                if self.db.delete_ativo(id):
                    messagebox.showinfo("Sucesso", "Ativo excluído com sucesso.")
                    self.load_data()
                else:
                    messagebox.showerror("Erro", "Não foi possível excluir o ativo.")
        else:
            messagebox.showwarning("Seleção", "Selecione um ativo para excluir.")

class EditWorkstationWindow(ctk.CTkToplevel):
    def __init__(self, master, db, hostname=None):
        super().__init__(master)
        
        self.db = db
        self.hostname = hostname
        self.is_edit = hostname is not None
        
        self.title("Editar Workstation" if self.is_edit else "Adicionar Workstation")
        self.geometry("800x600")
        self.center_window()
        self.state("zoomed")
        
        # Carregar dados se for edição
        self.data = None
        if self.is_edit:
            self.data = self.db.get_workstation(hostname)
        
        # Configurar interface
        self.setup_ui()
        
        # Preencher campos se for edição
        if self.is_edit and self.data:
            self.fill_fields()
        
        # Configurar protocolo de fechamento
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.grab_set()
    
    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def setup_ui(self):
        # Frame principal com scrollbar
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Canvas e scrollbar para formulário longo
        self.canvas = tk.Canvas(self.main_frame)
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ctk.CTkFrame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Campos do formulário
        self.fields = {}
        row = 0
        
        # Hostname (só leitura se for edição)
        self.fields['hostname'] = self.create_field("Hostname*", row, readonly=self.is_edit)
        row += 1
        
        # Usuario
        self.fields['usuario'] = self.create_field("Usuário", row)
        row += 1
        
        # Local*
        self.fields['local'] = self.create_combobox("Local*", row, COMBOBOX_WORKSTATIONS['Local'])
        row += 1
        
        # Sist. Oper.*
        self.fields['sist_oper'] = self.create_combobox("Sist. Oper.*", row, COMBOBOX_WORKSTATIONS['Sist. Oper.'])
        row += 1
        
        # Build*
        self.fields['build'] = self.create_combobox("Build*", row, COMBOBOX_WORKSTATIONS['Build'])
        row += 1

        # Bomgar*
        self.fields['bomgar'] = self.create_combobox("Bomgar*", row, COMBOBOX_WORKSTATIONS['Bomgar'])
        row += 1
        
        # Service Tag*
        self.fields['service_tag'] = self.create_field("Service Tag*", row)
        row += 1
        
        # MAC*
        self.fields['mac'] = self.create_field("MAC*", row)
        row += 1
        
        # Modelo*
        self.fields['modelo'] = self.create_combobox("Modelo*", row, COMBOBOX_WORKSTATIONS['Modelo'])
        row += 1
        
        # RAM*
        self.fields['ram'] = self.create_combobox("RAM*", row, COMBOBOX_WORKSTATIONS['RAM'])
        row += 1
        
        # Barramento
        self.fields['barramento'] = self.create_combobox("Barramento", row, COMBOBOX_WORKSTATIONS['Barramento'])
        row += 1
        
        # SSD*
        self.fields['ssd'] = self.create_combobox("SSD*", row, COMBOBOX_WORKSTATIONS['SSD'])
        row += 1
        
        # Chipset*
        self.fields['chipset'] = self.create_combobox("Chipset*", row, COMBOBOX_WORKSTATIONS['Chipset'])
        row += 1
        
        # Data Compra*
        self.fields['data_compra'] = self.create_date_field("Data Compra*", row)
        row += 1
        
        # Botões
        button_frame = ctk.CTkFrame(self.scrollable_frame)
        button_frame.grid(row=row, column=0, columnspan=2, pady=20, sticky="ew")
        
        self.save_btn = ctk.CTkButton(button_frame, text="Salvar", command=self.save)
        self.save_btn.pack(side="left", padx=10)
        
        self.cancel_btn = ctk.CTkButton(button_frame, text="Cancelar", command=self.on_close, fg_color="gray")
        self.cancel_btn.pack(side="left", padx=10)
    
    def create_field(self, label, row, readonly=False):
        lbl = ctk.CTkLabel(self.scrollable_frame, text=label)
        lbl.grid(row=row, column=0, padx=10, pady=5, sticky="w")
        
        if readonly:
            entry = ctk.CTkEntry(self.scrollable_frame, state="readonly")
        else:
            entry = ctk.CTkEntry(self.scrollable_frame)
        
        entry.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        
        self.scrollable_frame.grid_columnconfigure(1, weight=1)
        return entry
    
    def create_combobox(self, label, row, values):
        lbl = ctk.CTkLabel(self.scrollable_frame, text=label)
        lbl.grid(row=row, column=0, padx=10, pady=5, sticky="w")
        
        combo = AutocompleteCombobox(self.scrollable_frame, values=values)
        combo.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        
        self.scrollable_frame.grid_columnconfigure(1, weight=1)
        return combo
    
    def create_date_field(self, label, row):
        lbl = ctk.CTkLabel(self.scrollable_frame, text=label)
        lbl.grid(row=row, column=0, padx=10, pady=5, sticky="w")
        
        entry = DateEntry(self.scrollable_frame)
        entry.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        
        self.scrollable_frame.grid_columnconfigure(1, weight=1)
        return entry
    
    def fill_fields(self):
        self.fields['hostname'].configure(state="normal")
        self.fields['hostname'].insert(0, self.data['hostname'])
        if self.is_edit:
            self.fields['hostname'].configure(state="readonly")
        
        self.fields['usuario'].insert(0, self.data['usuario'] or "")
        self.fields['local'].set(self.data['local'])
        self.fields['sist_oper'].set(self.data['sist_oper'])
        self.fields['build'].set(self.data['build'])
        self.fields['bomgar'].set(self.data['bomgar'])
        self.fields['service_tag'].insert(0, self.data['service_tag'])
        self.fields['mac'].insert(0, self.data['mac'])
        self.fields['modelo'].set(self.data['modelo'])
        self.fields['ram'].set(self.data['ram'])
        self.fields['barramento'].set(self.data['barramento'] or "")
        self.fields['ssd'].set(self.data['ssd'])
        self.fields['chipset'].set(self.data['chipset'])
        self.fields['data_compra'].insert(0, self.data['data_compra'])
    
    def save(self):
        # Validar campos obrigatórios
        required_fields = {
            'hostname': 'Hostname',
            'local': 'Local',
            'sist_oper': 'Sistema Operacional',
            'build': 'Build',
            'bomgar': 'Bomgar',
            'service_tag': 'Service Tag',
            'mac': 'MAC',
            'modelo': 'Modelo',
            'ram': 'RAM',
            'ssd': 'SSD',
            'chipset': 'Chipset',
            'data_compra': 'Data de Compra'
        }
        
        data = {}
        for field, name in required_fields.items():
            value = self.fields[field].get().strip()
            if not value:
                messagebox.showerror("Erro", f"O campo {name} é obrigatório.")
                return
            data[field] = value
        
        # Campos opcionais
        data['usuario'] = self.fields['usuario'].get().strip()
        data['barramento'] = self.fields['barramento'].get().strip()
        
        # Validar data
        try:
            datetime.strptime(data['data_compra'], "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Erro", "Data de compra inválida. Use o formato DD/MM/YYYY.")
            return
        
        # Salvar no banco
        if self.is_edit:
            success = self.db.update_workstation(self.hostname, data)
        else:
            success = self.db.insert_workstation(data)
        
        if success:
            messagebox.showinfo("Sucesso", "Workstation salva com sucesso.")
            self.master.load_data()
            self.on_close()
        else:
            if self.is_edit:
                messagebox.showerror("Erro", "Não foi possível atualizar a workstation.")
            else:
                messagebox.showerror("Erro", "Não foi possível adicionar a workstation. Hostname já existe.")
    
    def on_close(self):
        self.grab_release()
        self.destroy()

class EditAtivoWindow(ctk.CTkToplevel):
    def __init__(self, master, db, id=None):
        super().__init__(master)
        
        self.db = db
        self.id = id
        self.is_edit = id is not None
        
        self.title("Editar Ativo" if self.is_edit else "Adicionar Ativo")
        self.geometry("600x500")
        self.center_window()
        self.state("zoomed")
        
        # Carregar dados se for edição
        self.data = None
        if self.is_edit:
            self.data = self.db.get_ativo(id)
        
        # Configurar interface
        self.setup_ui()
        
        # Preencher campos se for edição
        if self.is_edit and self.data:
            self.fill_fields()
        
        # Configurar protocolo de fechamento
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.grab_set()
    
    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def setup_ui(self):
        # Frame principal
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Campos do formulário
        self.fields = {}
        row = 0
        
        # Tipo*
        self.fields['tipo'] = self.create_combobox("Tipo*", row, COMBOBOX_ATIVOS['Tipo'])
        row += 1
        
        # Marca*
        self.fields['marca'] = self.create_combobox("Marca*", row, COMBOBOX_ATIVOS['Marca'])
        row += 1
        
        # Modelo*
        self.fields['modelo'] = self.create_field("Modelo*", row)
        row += 1
        
        # Estado*
        self.fields['estado'] = self.create_combobox("Estado*", row, COMBOBOX_ATIVOS['Estado'])
        row += 1
        
        # Notas
        lbl = ctk.CTkLabel(self.main_frame, text="Notas")
        lbl.grid(row=row, column=0, padx=10, pady=5, sticky="nw")
        
        self.fields['notas'] = ctk.CTkTextbox(self.main_frame, height=80)
        self.fields['notas'].grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        row += 1
        
        # Data Compra
        self.fields['data_compra'] = self.create_date_field("Data Compra", row)
        row += 1
        
        # Máquina Associada
        # Primeiro, obter lista de workstations
        workstations = self.db.get_workstations()
        workstation_names = [ws['hostname'] for ws in workstations]
        
        self.fields['maquina_associada'] = self.create_combobox("Máquina Associada", row, workstation_names)
        row += 1
        
        # Botões
        button_frame = ctk.CTkFrame(self.main_frame)
        button_frame.grid(row=row, column=0, columnspan=2, pady=20, sticky="ew")
        
        self.save_btn = ctk.CTkButton(button_frame, text="Salvar", command=self.save)
        self.save_btn.pack(side="left", padx=10)
        
        self.cancel_btn = ctk.CTkButton(button_frame, text="Cancelar", command=self.on_close, fg_color="gray")
        self.cancel_btn.pack(side="left", padx=10)
        
        self.main_frame.grid_columnconfigure(1, weight=1)
    
    def create_field(self, label, row):
        lbl = ctk.CTkLabel(self.main_frame, text=label)
        lbl.grid(row=row, column=0, padx=10, pady=5, sticky="w")
        
        entry = ctk.CTkEntry(self.main_frame)
        entry.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        
        return entry
    
    def create_combobox(self, label, row, values):
        lbl = ctk.CTkLabel(self.main_frame, text=label)
        lbl.grid(row=row, column=0, padx=10, pady=5, sticky="w")
        
        combo = AutocompleteCombobox(self.main_frame, values=values)
        combo.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        
        return combo
    
    def create_date_field(self, label, row):
        lbl = ctk.CTkLabel(self.main_frame, text=label)
        lbl.grid(row=row, column=0, padx=10, pady=5, sticky="w")
        
        entry = DateEntry(self.main_frame)
        entry.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        
        return entry
    
    def fill_fields(self):
        self.fields['tipo'].set(self.data['tipo'])
        self.fields['marca'].set(self.data['marca'])
        self.fields['modelo'].insert(0, self.data['modelo'])
        self.fields['estado'].set(self.data['estado'])
        self.fields['notas'].insert("1.0", self.data['notas'] or "")
        self.fields['data_compra'].insert(0, self.data['data_compra'] or "")
        self.fields['maquina_associada'].set(self.data['maquina_associada'] or "")
    
    def save(self):
        # Validar campos obrigatórios
        required_fields = {
            'tipo': 'Tipo',
            'marca': 'Marca',
            'modelo': 'Modelo',
            'estado': 'Estado'
        }
        
        data = {}
        for field, name in required_fields.items():
            value = self.fields[field].get().strip()
            if not value:
                messagebox.showerror("Erro", f"O campo {name} é obrigatório.")
                return
            data[field] = value
        
        # Campos opcionais
        data['notas'] = self.fields['notas'].get("1.0", "end-1c").strip()
        data['data_compra'] = self.fields['data_compra'].get().strip()
        data['maquina_associada'] = self.fields['maquina_associada'].get().strip()
        
        # Validar data se fornecida
        if data['data_compra']:
            try:
                datetime.strptime(data['data_compra'], "%d/%m/%Y")
            except ValueError:
                messagebox.showerror("Erro", "Data de compra inválida. Use o formato DD/MM/YYYY.")
                return
        
        # Salvar no banco
        if self.is_edit:
            success = self.db.update_ativo(self.id, data)
        else:
            success = self.db.insert_ativo(data)
        
        if success:
            messagebox.showinfo("Sucesso", "Ativo salvo com sucesso.")
            self.master.load_data()
            self.on_close()
        else:
            messagebox.showerror("Erro", "Não foi possível salvar o ativo.")
    
    def on_close(self):
        self.grab_release()
        self.destroy()

class DuplicateWorkstationWindow(ctk.CTkToplevel):
    def __init__(self, master, db, hostname):
        super().__init__(master)
        
        self.db = db
        self.original_hostname = hostname
        self.original_data = self.db.get_workstation(hostname)
        
        self.title("Duplicar Workstation")
        self.geometry("400x200")
        self.center_window()
        self.state("zoomed")
        
        # Configurar interface
        self.setup_ui()
        
        # Configurar protocolo de fechamento
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.grab_set()
    
    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def setup_ui(self):
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        ctk.CTkLabel(self.main_frame, text="Duplicar Workstation", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        ctk.CTkLabel(self.main_frame, text=f"Original: {self.original_hostname}").pack(pady=5)
        
        # Hostname
        hostname_frame = ctk.CTkFrame(self.main_frame)
        hostname_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(hostname_frame, text="Novo Hostname*").pack(side="left")
        self.hostname_entry = ctk.CTkEntry(hostname_frame)
        self.hostname_entry.pack(side="right", fill="x", expand=True, padx=(10, 0))
        
        # MAC
        mac_frame = ctk.CTkFrame(self.main_frame)
        mac_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(mac_frame, text="Novo MAC*").pack(side="left")
        self.mac_entry = ctk.CTkEntry(mac_frame)
        self.mac_entry.pack(side="right", fill="x", expand=True, padx=(10, 0))
        
        # Botões
        button_frame = ctk.CTkFrame(self.main_frame)
        button_frame.pack(pady=10)
        
        self.duplicate_btn = ctk.CTkButton(button_frame, text="Duplicar", command=self.duplicate)
        self.duplicate_btn.pack(side="left", padx=10)
        
        self.cancel_btn = ctk.CTkButton(button_frame, text="Cancelar", command=self.on_close, fg_color="gray")
        self.cancel_btn.pack(side="left", padx=10)
    
    def duplicate(self):
        new_hostname = self.hostname_entry.get().strip()
        new_mac = self.mac_entry.get().strip()
        
        if not new_hostname:
            messagebox.showerror("Erro", "O campo Hostname é obrigatório.")
            return
        
        if not new_mac:
            messagebox.showerror("Erro", "O campo MAC é obrigatório.")
            return
        
        # Criar cópia dos dados originais
        data = self.original_data.copy()
        data['hostname'] = new_hostname
        data['mac'] = new_mac
        
        # Inserir nova workstation
        if self.db.insert_workstation(data):
            messagebox.showinfo("Sucesso", "Workstation duplicada com sucesso.")
            self.master.load_data()
            self.on_close()
        else:
            messagebox.showerror("Erro", "Não foi possível duplicar a workstation. Hostname já existe.")
    
    def on_close(self):
        self.grab_release()
        self.destroy()

class DetailsWindow(ctk.CTkToplevel):
    def __init__(self, master, db, hostname):
        super().__init__(master)
        
        self.db = db
        self.hostname = hostname
        self.workstation = self.db.get_workstation(hostname)
        self.ativos = self.db.get_ativos(maquina_associada=hostname)
        
        self.title(f"Detalhes da Workstation: {hostname}")
        self.geometry("1000x700")
        self.center_window()
        self.state("zoomed")
        
        # Configurar interface
        self.setup_ui()
        
        # Configurar protocolo de fechamento
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.grab_set()
    
    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def setup_ui(self):
        # Frame principal com abas
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Aba de detalhes da workstation
        self.details_tab = self.tabview.add("Detalhes da Workstation")
        
        # Aba de ativos associados
        self.assets_tab = self.tabview.add("Ativos Associados")
        
        # Configurar aba de detalhes
        self.setup_details_tab()
        
        # Configurar aba de ativos
        self.setup_assets_tab()
        
        # Botão fechar
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(fill="x", padx=10, pady=10)
        
        self.close_btn = ctk.CTkButton(button_frame, text="Fechar", command=self.on_close, fg_color="#F1F1F1", hover_color="#DADADA", border_color="#7C7C7C", border_width=1, text_color="black")
        self.close_btn.pack(side="right")
    
    def setup_details_tab(self):
        # Frame com scrollbar para detalhes
        details_frame = ctk.CTkFrame(self.details_tab, fg_color="transparent")
        details_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Canvas e scrollbar
        canvas = tk.Canvas(details_frame)
        scrollbar = ttk.Scrollbar(details_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ctk.CTkFrame(canvas, fg_color="transparent")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Exibir detalhes da workstation
        if self.workstation:
            fields = [
                ("Hostname", self.workstation['hostname']),
                ("Usuário", self.workstation['usuario'] or "N/A"),
                ("Local", self.workstation['local']),
                ("Sistema Operacional", self.workstation['sist_oper']),
                ("Build", self.workstation['build']),
                ("Bomgar", self.workstation['bomgar']),
                ("Service Tag", self.workstation['service_tag']),
                ("MAC", self.workstation['mac']),
                ("Modelo", self.workstation['modelo']),
                ("RAM", self.workstation['ram']),
                ("Barramento", self.workstation['barramento'] or "N/A"),
                ("SSD", self.workstation['ssd']),
                ("Chipset", self.workstation['chipset']),
                ("Data de Compra", self.workstation['data_compra']),
                ("Tempo de Uso", self.workstation['tempo_uso'])
            ]
            
            for i, (label, value) in enumerate(fields):
                lbl = ctk.CTkLabel(scrollable_frame, text=label + ":", font=ctk.CTkFont(weight="bold"))
                lbl.grid(row=i, column=0, padx=10, pady=5, sticky="w")
                
                value_entry = ctk.CTkEntry(scrollable_frame, width=400)
                value_entry.insert(0, value)
                value_entry.configure(state="readonly")
                value_entry.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
            
            scrollable_frame.grid_columnconfigure(1, weight=1)
    
    def setup_assets_tab(self):
        # Frame para ativos
        assets_frame = ctk.CTkFrame(self.assets_tab)
        assets_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        if self.ativos:
            # Treeview para ativos
            tree_frame = ctk.CTkFrame(assets_frame)
            tree_frame.pack(fill="both", expand=True)
            
            tree = ttk.Treeview(tree_frame, columns=(
                "tipo", "marca", "modelo", "estado", "data_compra", "tempo_uso"
            ), show="headings", height=15, selectmode="extended")
            
            # Definir colunas
            columns = {
                "tipo": "Tipo",
                "marca": "Marca",
                "modelo": "Modelo",
                "estado": "Estado",
                "data_compra": "Data Compra",
                "tempo_uso": "Tempo de Uso"
            }
            
            for col, text in columns.items():
                tree.heading(col, text=text)
                tree.column(col, width=100, minwidth=50)
            
            # Scrollbars
            v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
            
            # Layout
            tree.grid(row=0, column=0, sticky="nsew")
            v_scrollbar.grid(row=0, column=1, sticky="ns")
            h_scrollbar.grid(row=1, column=0, sticky="ew")
            
            tree_frame.grid_rowconfigure(0, weight=1)
            tree_frame.grid_columnconfigure(0, weight=1)
            
            # Preencher com dados
            for ativo in self.ativos:
                tree.insert("", "end", values=(
                    ativo['tipo'], ativo['marca'], ativo['modelo'], 
                    ativo['estado'], ativo['data_compra'] or "N/A", ativo['tempo_uso']
                ))
        else:
            ctk.CTkLabel(assets_frame, text="Nenhum ativo associado a esta workstation.").pack(expand=True)
    
    def on_close(self):
        self.grab_release()
        self.destroy()

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()