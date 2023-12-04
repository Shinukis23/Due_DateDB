from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkinter import messagebox
from tkinter import colorchooser
from configparser import ConfigParser
import oracledb
from screeninfo import get_monitors
import webbrowser

user = "ADMIN"
dsn = "tapreturnsdb_high"

monitors = get_monitors()
screen = monitors[0]

pw = "pROCESOS2023"
wallet_pw = "!Tr3svecestres!"
parser = ConfigParser()
parser.read("treebase.ini")
saved_primary_color = parser.get('colors', 'primary_color')
saved_secondary_color = parser.get('colors', 'secondary_color')
saved_highlight_color = parser.get('colors', 'highlight_color')
def main_function(usuario, main_window, records, opcion, hoy):
    main_window.attributes('-disabled', True)
    status_window = tk.Toplevel(main_window)git
    status_window.title('Autorizacion de regresos - Servicio al Cliente')
    status_window.iconbitmap("logoicon.ico")
    status_window.geometry("1100x600")
    status_window.attributes('-topmost', True)

    def validation_dt(i, text, new_text):
        return len(new_text) == 0 or len(new_text) < 10 and text.isdecimal()

    def update_button_state():
        if DT.get() and validation_dt(1, DT.get(), ""):
            update_button.config(state="normal")
        else:
            update_button.config(state="disabled")

    def query_database(usuario1):

        for record in my_tree.get_children():
            my_tree.delete(record)

        global count
        count = 0
        if records:

            for record in records:
                if count % 2 == 0:
                    my_tree.insert(parent='', index='end', iid=count, text='', values=(
                        record[0], record[11], record[12], record[14], record[10], record[1], record[2], record[3],
                        record[4], record[5], record[6], record[13], record[7], record[9], record[8], record[17],
                        record[15]),
                                   tags=('evenrow',))
                else:
                    my_tree.insert(parent='', index='end', iid=count, text='', values=(
                        record[0], record[11], record[12], record[14], record[10], record[1], record[2], record[3],
                        record[4], record[5], record[6], record[13], record[7], record[9], record[8], record[17],
                        record[15]),
                                   tags=('oddrow',))
                count += 1

    def search_records():
        lookup_record = search_entry.get()
        search.destroy()

        for record in my_tree.get_children():
            my_tree.delete(record)

        c = con.cursor()

        c.execute("SELECT rowid, * FROM Regresos WHERE last_name like ?", (lookup_record,))
        records = c.fetchall()

        global count
        count = 0

        for record in records:
            if count % 2 == 0:
                my_tree.insert(parent='', index='end', iid=count, text='', values=(
                    record[0], record[11], record[12], record[14], record[10], record[1], record[2], record[3],
                    record[4],
                    record[5], record[6], record[13], record[7], record[9], record[8], record[17], record[15]),
                               tags=('evenrow',))

            else:
                my_tree.insert(parent='', index='end', iid=count, text='', values=(
                    record[0], record[11], record[12], record[14], record[10], record[1], record[2], record[3],
                    record[4],
                    record[5], record[6], record[13], record[7], record[9], record[8], record[17], record[15]),
                               tags=('oddrow',))
            count += 1

        con.commit()
        con.close()

    def lookup_records():
        global search_entry, search

        search = Toplevel(status_window)
        search.title("Lookup Records")
        search.geometry("400x200")
        search.iconbitmap('logoicon.ico')

        search_frame = LabelFrame(search, text="Last Name")
        search_frame.pack(padx=10, pady=10)

        search_entry = Entry(search_frame, font=("Helvetica", 18))
        search_entry.pack(pady=20, padx=20)

        search_button = Button(search, text="Search Records", command=search_records)
        search_button.pack(padx=20, pady=20)

    def primary_color():
        # Pick Color
        primary_color = colorchooser.askcolor()[1]

        # Update Treeview Color
        if primary_color:
            # Create Striped Row Tags
            my_tree.tag_configure('evenrow', background=primary_color)

            # Config file
            parser = ConfigParser()
            parser.read("treebase.ini")
            # Set the color change
            parser.set('colors', 'primary_color', primary_color)
            # Save the config file
            with open('treebase.ini', 'w') as configfile:
                parser.write(configfile)

    def secondary_color():
        # Pick Color
        secondary_color = colorchooser.askcolor()[1]

        # Update Treeview Color
        if secondary_color:
            # Create Striped Row Tags
            my_tree.tag_configure('oddrow', background=secondary_color)

            # Config file
            parser = ConfigParser()
            parser.read("treebase.ini")
            # Set the color change
            parser.set('colors', 'secondary_color', secondary_color)
            # Save the config file
            with open('treebase.ini', 'w') as configfile:
                parser.write(configfile)

    def highlight_color():
        # Pick Color
        highlight_color = colorchooser.askcolor()[1]

        # Update Treeview Color
        # Change Selected Color
        if highlight_color:
            style.map('Treeview',
                      background=[('selected', highlight_color)])

            # Config file
            parser = ConfigParser()
            parser.read("treebase.ini")
            # Set the color change
            parser.set('colors', 'highlight_color', highlight_color)
            # Save the config file
            with open('treebase.ini', 'w') as configfile:
                parser.write(configfile)

    def query_database1():
        query_database(usuario)

    def open_link_in_browser(link):
        webbrowser.open(link)

    def reset_colors():
        # Save original colors to config file
        parser = ConfigParser()
        parser.read('treebase.ini')
        parser.set('colors', 'primary_color', 'lightgray')
        parser.set('colors', 'secondary_color', 'white')
        parser.set('colors', 'highlight_color', '#347083')
        with open('treebase.ini', 'w') as configfile:
            parser.write(configfile)
        my_tree.tag_configure('oddrow', background='white')
        my_tree.tag_configure('evenrow', background='lightblue')
        style.map('Treeview',
                  background=[('selected', '#347083')])

    my_menu = Menu(status_window)
    status_window.config(menu=my_menu)
    help_menu = Menu(my_menu, tearoff=0)
    my_menu.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="Due_date Help", command=lookup_records)
    help_menu.add_separator()
    help_menu.add_command(label="About Due_date", command=query_database1)
    '''
	for record in data:
		c.execute("INSERT INTO Regresos VALUES (:first_name, :last_name, :id, :address, :city, :state, :zipcode)", 
			{
			'first_name': record[0],
			'last_name': record[1],
			'id': record[2],
			'address': record[3],
			'city': record[4],
			'state': record[5],
			'zipcode': record[6]
			}
			)
	'''
    style = ttk.Style()

    style.theme_use('default')
    style.configure("Treeview",
                    background="#D3D3D3",
                    foreground="black",
                    rowheight=20,
                    fieldbackground="#D3D3D3")
    style.map('Treeview',
              background=[('selected', saved_highlight_color)])
    tree_frame = Frame(status_window)
    tree_frame.pack(pady=10)

    y_scroll = ttk.Scrollbar(tree_frame, orient=VERTICAL)
    x_scroll = ttk.Scrollbar(tree_frame, orient=HORIZONTAL)
    x_scroll.pack(side=BOTTOM, fill=X)
    y_scroll.pack(side=RIGHT, fill=Y)
    my_tree = ttk.Treeview(tree_frame, yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
    my_tree.pack()
    my_tree.configure(selectmode="extended")
    y_scroll.config(command=my_tree.yview)
    x_scroll.config(command=my_tree.xview)

    my_tree['columns'] = (
        "ID", "Status", "Usuario", "Fecha de requerimiento", "Delivery Ticket", "Factura", "Fecha venta",
        "Tipo de parte",
        "# Stock", "Motivo Regreso", "Descripcion motivo", "Vendedor", "Autorizado", "Autorizado por",
        "Fecha", "Notas")

    # Format Our Columns
    my_tree.column("#0", width=0, stretch=NO)
    my_tree.column("ID", anchor=CENTER, width=70)
    my_tree.column("Status", anchor=CENTER, width=100)
    my_tree.column("Usuario", anchor=W, width=140)
    my_tree.column("Fecha de requerimiento", anchor=W, width=140)
    my_tree.column("Factura", anchor=CENTER, width=70)
    my_tree.column("Fecha venta", anchor=CENTER, width=140)
    my_tree.column("Tipo de parte", anchor=CENTER, width=140)
    my_tree.column("# Stock", anchor=CENTER, width=100)
    my_tree.column("Motivo Regreso", anchor=CENTER, width=160)
    my_tree.column("Descripcion motivo", anchor=CENTER, width=600)
    my_tree.column("Vendedor", anchor=CENTER, width=140)
    my_tree.column("Autorizado", anchor=CENTER, width=100)
    my_tree.column("Autorizado por", anchor=CENTER, width=100)
    my_tree.column("Fecha", anchor=CENTER, width=140)
    my_tree.column("Delivery Ticket", anchor=CENTER, width=100)
    my_tree.column("Notas", anchor=CENTER, width=500)

    my_tree.heading("#0", text="", anchor=W)
    my_tree.heading("ID", text="ID", anchor=CENTER)
    my_tree.heading("Status", text="Status", anchor=CENTER)
    my_tree.heading("Usuario", text="Usuario", anchor=W)
    my_tree.heading("Fecha de requerimiento", text="Fecha de requerimiento", anchor=W)
    my_tree.heading("Factura", text="Factura", anchor=CENTER)
    my_tree.heading("Fecha venta", text="Fecha venta", anchor=CENTER)
    my_tree.heading("Tipo de parte", text="Tipo de parte", anchor=CENTER)
    my_tree.heading("# Stock", text="# Stock", anchor=CENTER)
    my_tree.heading("Motivo Regreso", text="Motivo Regreso", anchor=CENTER)
    my_tree.heading("Descripcion motivo", text="Descripcion motivo", anchor=CENTER)
    my_tree.heading("Vendedor", text="Vendedor", anchor=CENTER)
    my_tree.heading("Autorizado", text="Autorizado", anchor=CENTER)
    my_tree.heading("Autorizado por", text="Autorizado por", anchor=CENTER)
    my_tree.heading("Fecha", text="Fecha", anchor=CENTER)
    my_tree.heading("Delivery Ticket", text="Delivery Ticket", anchor=CENTER)
    my_tree.heading("Notas", text="Notas", anchor=CENTER)

    my_tree.tag_configure('oddrow', background=saved_secondary_color)
    my_tree.tag_configure('evenrow', background=saved_primary_color)

    data_frame = LabelFrame(status_window, text="Registro")
    data_frame.pack(fill="x", expand="yes", padx=20)

    fac_label = Label(data_frame, text="Factura")
    fac_label.grid(row=0, column=0, padx=10, pady=10)
    fac_entry = Entry(data_frame)
    fac_entry.grid(row=0, column=1, padx=10, pady=10)

    Fv_label = Label(data_frame, text="Fecha Venta")
    Fv_label.grid(row=1, column=0, padx=10, pady=10)
    Fv_entry = Entry(data_frame)
    Fv_entry.grid(row=1, column=1, padx=10, pady=10)

    Sell_label = Label(data_frame, text="Vendedor")
    Sell_label.grid(row=2, column=0, padx=10, pady=10)
    Sell_entry = Entry(data_frame)
    Sell_entry.grid(row=2, column=1, padx=10, pady=10)

    Stck_label = Label(data_frame, text="# Stock")
    Stck_label.grid(row=3, column=0, padx=10, pady=10)
    Stck_entry = Entry(data_frame)
    Stck_entry.grid(row=3, column=1, padx=10, pady=10)

    User_label = Label(data_frame, text="Requerido por")
    User_label.grid(row=0, column=3, padx=10, pady=10)
    User_entry = Entry(data_frame)
    User_entry.grid(row=0, column=4, padx=10, pady=10)

    PartCode_label = Label(data_frame, text="Tipo de Parte")
    PartCode_label.grid(row=1, column=3, padx=10, pady=10)
    PartCode_entry = Entry(data_frame)
    PartCode_entry.grid(row=1, column=4, padx=10, pady=10)

    Autby_label = Label(data_frame, text="Autorizado por")
    Autby_label.grid(row=2, column=3, padx=10, pady=10)
    Autby_entry = Entry(data_frame)
    Autby_entry.grid(row=2, column=4, padx=10, pady=10)

    FechaAut_label = Label(data_frame, text="Fecha Autorizacion")
    FechaAut_label.grid(row=3, column=3, padx=10, pady=10)
    FechaAut_entry = Entry(data_frame)
    FechaAut_entry.grid(row=3, column=4, padx=10, pady=10)

    Motivo_label = Label(data_frame, text="Motivo Regreso")
    Motivo_label.grid(row=0, column=6, padx=10, pady=10)
    Motivo_entry = Entry(data_frame)
    Motivo_entry.grid(row=0, column=7, padx=10, pady=10)

    Desmotivo_label = Label(data_frame, text="Descripcion motivo")
    Desmotivo_label.grid(row=1, column=6, padx=10, pady=10)
    Desmotivo_entry = Entry(data_frame)
    Desmotivo_entry = tk.Text(data_frame, wrap="word", width=30, height=6)
    Desmotivo_entry.grid(row=2, column=6, rowspan=2, columnspan=2, padx=10, pady=10)

    Notas_label = Label(data_frame, text="Notas")
    Notas_label.grid(row=1, column=8, padx=10, pady=10)
    Notas_entry = Entry(data_frame)
    Notas_entry = tk.Text(data_frame, wrap="word", width=30, height=6)
    Notas_entry.grid(row=2, column=8, rowspan=2, columnspan=2, padx=10, pady=10)

    def up():
        rows = my_tree.selection()
        for row in rows:
            my_tree.move(row, my_tree.parent(row), my_tree.index(row) - 1)

    def down():
        rows = my_tree.selection()
        for row in reversed(rows):
            my_tree.move(row, my_tree.parent(row), my_tree.index(row) + 1)

    def remove_one():
        x = my_tree.selection()[0]
        my_tree.delete(x)
        conn = oracledb.connect(user=user, password=pw, dsn=dsn, config_dir=r"C:\\network\\admin",
                                wallet_location=r"C:\\network\\admin", wallet_password=wallet_pw)
        c = conn.cursor()
        c.execute("DELETE from Regresos WHERE oid=" + id_entry.get())
        conn.commit()
        conn.close()

        clear_entries()
        messagebox.showinfo("Deleted!", "El registro ha sido eliminado")

    def remove_many():
        response = messagebox.askyesno("Precaucion!!!!",
                                       "Esto borrara todos los registros\nEstas seguro?!")

        if response == 1:
            x = my_tree.selection()
            ids_to_delete = []
            for record in x:
                ids_to_delete.append(my_tree.item(record, 'values')[2])

            for record in x:
                my_tree.delete(record)

            conn = oracledb.connect(user=user, password=pw, dsn=dsn, config_dir=r"C:\\network\\admin",
                                    wallet_location=r"C:\\network\\admin", wallet_password=wallet_pw)

            c = conn.cursor()
            c.executemany("DELETE FROM Regresos WHERE id = ?", [(a,) for a in ids_to_delete])

            ids_to_delete = []

            conn.commit()
            conn.close()
            clear_entries()

    def close_window2():
        reset_colors()
        style.theme_use('vista')
        status_window.destroy()
        main_window.attributes('-disabled', False)

    def rechazar():
        selected = my_tree.focus()
        result1 = hoy.strftime("%d-%b-%y")
        current_values = my_tree.item(selected, 'values')
        current_values = (
            current_values[0],
            "Done",
            User_entry.get(),
            current_values[3],
            dt_entry.get(),
            fac_entry.get(),
            Fv_entry.get(),
            PartCode_entry.get(),
            Stck_entry.get(),
            Motivo_entry.get(),
            Desmotivo_entry.get("1.0", END),
            Sell_entry.get(),
            current_values[12],
            Autby_entry.get(),
            FechaAut_entry.get(),
            Notas_entry.get("1.0", END)
        )
        my_tree.item(selected, values=current_values)
        conn = oracledb.connect(user=user, password=pw, dsn=dsn, config_dir=r"C:\\network\\admin",
                                wallet_location=r"C:\\network\\admin", wallet_password=wallet_pw)
        c = conn.cursor()

        c.execute("""UPDATE Regresos SET
                                estatus_regreso = : estatus,
                                autorizado = : resul,
                    			persona_autorizacion = :persona,
                    			fecha_autorizacion = : fechareject,
                    			observacion = : motivodesc,
                    			descripcion_motivo = : descrmotivo

        			WHERE ID = :ID""",
                  {
                      'estatus': "DoneNG",
                      'resul': "No",
                      'persona': usuario,
                      'fechareject': result1,
                      'motivodesc': Notas_entry.get("1.0", END),
                      'descrmotivo': Desmotivo_entry.get("1.0", END),
                      'ID': current_values[0],
                  })

        conn.commit()
        conn.close()

        habilitar(1)
        fac_entry.delete(0, END)
        Sell_entry.delete(0, END)
        User_entry.delete(0, END)
        Fv_entry.delete(0, END)
        Stck_entry.delete(0, END)
        PartCode_entry.delete(0, END)
        Motivo_entry.delete(0, END)
        Autby_entry.delete(0, END)
        FechaAut_entry.delete(0, END)
        Desmotivo_entry.delete("1.0", END)
        Notas_entry.delete("1.0", END)
        dt_entry.delete(0, END)
        habilitar(0)
        my_tree.delete(selected)

    def habilitar(enable):
        if enable == 1:
            estado = tk.NORMAL
        else:
            estado = tk.DISABLED
        fac_entry.config(state=estado)
        Sell_entry.config(state=estado)
        User_entry.config(state=estado)
        Fv_entry.config(state=estado)
        Stck_entry.config(state=estado)
        PartCode_entry.config(state=estado)
        Motivo_entry.config(state=estado)
        Autby_entry.config(state=estado)
        FechaAut_entry.config(state=estado)
        Desmotivo_entry.config(state=estado)
        Notas_entry.config(state=estado)
        update_button.config(state=estado)
        if opcion == 0:
            dt_entry.config(state=estado)
        elif opcion != 0 and usuario in ["MARIANA MENDOZA", "JUAN ORTIZ", "PEDRO JUVERA", "MANUEL RAZO",
                                         "EMMANUEL LOPEZ"]:
            dt_entry.config(state=tk.DISABLED)
            update_button.config(state=tk.DISABLED)
        else:
            dt_entry.config(state=tk.DISABLED)

    def clear_entries():
        habilitar(1)
        fac_entry.delete(0, END)
        Sell_entry.delete(0, END)
        User_entry.delete(0, END)
        Fv_entry.delete(0, END)
        Stck_entry.delete(0, END)
        PartCode_entry.delete(0, END)
        Motivo_entry.delete(0, END)
        Autby_entry.delete(0, END)
        FechaAut_entry.delete(0, END)
        Desmotivo_entry.delete("1.0", END)
        Notas_entry.delete("1.0", END)
        dt_entry.delete(0, END)
        habilitar(0)

    def select_record(e):
        print("opcion: ", opcion)
        habilitar(1)
        fac_entry.delete(0, END)
        Sell_entry.delete(0, END)
        User_entry.delete(0, END)
        Fv_entry.delete(0, END)
        Stck_entry.delete(0, END)
        PartCode_entry.delete(0, END)
        Motivo_entry.delete(0, END)
        Autby_entry.delete(0, END)
        FechaAut_entry.delete(0, END)
        Desmotivo_entry.delete("1.0", END)
        Notas_entry.delete("1.0", END)
        dt_entry.delete(0, END)

        selected = my_tree.focus()
        values = my_tree.item(selected, 'values')
        column = my_tree.identify_column(e.x)
        fac_entry.insert(0, values[5])
        Sell_entry.insert(0, values[11])
        User_entry.insert(0, values[2])
        Fv_entry.insert(0, values[6])
        Stck_entry.insert(0, values[8])
        Autby_entry.insert(0, values[13])
        PartCode_entry.insert(0, values[7])
        FechaAut_entry.insert(0, values[14])
        Desmotivo_entry.insert("1.0", values[10])
        Motivo_entry.insert(0, values[9])
        Notas_entry.insert("1.0", values[15])
        if opcion == 0 and usuario not in ["MARIANA MENDOZA", "JUAN ORTIZ", "PEDRO JUVERA", "MANUEL RAZO",
                                           "EMMANUEL LOPEZ"]:

            habilitar(0)
            dt_entry.config(state=tk.NORMAL)
            update_button.config(state=tk.NORMAL)
        elif opcion == 0 and usuario in ["MARIANA MENDOZA", "JUAN ORTIZ", "PEDRO JUVERA", "MANUEL RAZO",
                                         "EMMANUEL LOPEZ"]:
            dt_entry.config(state=tk.NORMAL)
            habilitar(1)
        elif opcion != 0 and usuario not in ["MARIANA MENDOZA", "JUAN ORTIZ", "PEDRO JUVERA", "MANUEL RAZO",
                                             "EMMANUEL LOPEZ"]:
            dt_entry.config(state=tk.DISABLED)
            habilitar(0)
        elif opcion != 0 and usuario in ["MARIANA MENDOZA", "JUAN ORTIZ", "PEDRO JUVERA", "MANUEL RAZO",
                                         "EMMANUEL LOPEZ"]:
            dt_entry.config(state=tk.DISABLED)
            add_button.config(state=tk.NORMAL)
            remove_all_button.config(state=tk.NORMAL)
            if selected and column == '#10':
                link = my_tree.item(selected, 'values')[16]
                if link != "None":
                    open_link_in_browser(link)
            habilitar(1)

    def update_record():
        selected = my_tree.focus()
        current_values = my_tree.item(selected, 'values')
        current_values = (
            current_values[0],
            "Done",
            User_entry.get(),
            current_values[3],
            dt_entry.get(),
            fac_entry.get(),
            Fv_entry.get(),
            PartCode_entry.get(),
            Stck_entry.get(),
            Motivo_entry.get(),
            Desmotivo_entry.get("1.0", END),
            Sell_entry.get(),
            current_values[12],
            Autby_entry.get(),
            FechaAut_entry.get(),
            Notas_entry.get("1.0", END)
        )
        my_tree.item(selected, values=current_values)
        conn = oracledb.connect(user=user, password=pw, dsn=dsn, config_dir=r"C:\\network\\admin",
                                wallet_location=r"C:\\network\\admin", wallet_password=wallet_pw)

        c = conn.cursor()

        c.execute("""UPDATE Regresos SET
                        estatus_regreso = : estatus,
            			num_orden_regreso = :dt

			WHERE ID = :ID""",
                  {
                      'estatus': "Done",
                      'dt': dt_entry.get(),
                      'ID': current_values[0],
                  })

        conn.commit()
        conn.close()
        habilitar(1)
        fac_entry.delete(0, END)
        Sell_entry.delete(0, END)
        User_entry.delete(0, END)
        Fv_entry.delete(0, END)
        Stck_entry.delete(0, END)
        PartCode_entry.delete(0, END)
        Motivo_entry.delete(0, END)
        Autby_entry.delete(0, END)
        FechaAut_entry.delete(0, END)
        Desmotivo_entry.delete("1.0", END)
        Notas_entry.delete("1.0", END)
        dt_entry.delete(0, END)
        habilitar(0)
        my_tree.delete(selected)

    def autorizar():
        selected = my_tree.focus()
        result1 = hoy.strftime("%d-%b-%y")
        current_values = my_tree.item(selected, 'values')
        current_values = (
            current_values[0],
            "Done",
            User_entry.get(),
            current_values[3],
            dt_entry.get(),
            fac_entry.get(),
            Fv_entry.get(),
            PartCode_entry.get(),
            Stck_entry.get(),
            Motivo_entry.get(),
            Desmotivo_entry.get("1.0", END),
            Sell_entry.get(),
            current_values[12],
            Autby_entry.get(),
            FechaAut_entry.get(),
            Notas_entry.get("1.0", END)
        )
        my_tree.item(selected, values=current_values)
        conn = oracledb.connect(user=user, password=pw, dsn=dsn, config_dir=r"C:\\network\\admin",
                                wallet_location=r"C:\\network\\admin", wallet_password=wallet_pw)
        c = conn.cursor()

        c.execute("""UPDATE Regresos SET
                                        estatus_regreso = : estatus,
                                        autorizado = : resul,
                            			persona_autorizacion = :persona,
                            			fecha_autorizacion = : fechareject,
                            			observacion = : motivodesc,
                            			descripcion_motivo = : descrmotivo

                			WHERE ID = :ID""",
                  {
                      'estatus': "PendingDT",
                      'resul': "Yes",
                      'persona': usuario,
                      'fechareject': result1,
                      'motivodesc': Notas_entry.get("1.0", END),
                      'descrmotivo': Desmotivo_entry.get("1.0", END),
                      'ID': current_values[0],
                  })

        conn.commit()
        conn.close()
        habilitar(1)
        fac_entry.delete(0, END)
        Sell_entry.delete(0, END)
        User_entry.delete(0, END)
        Fv_entry.delete(0, END)
        Stck_entry.delete(0, END)
        PartCode_entry.delete(0, END)
        Motivo_entry.delete(0, END)
        Autby_entry.delete(0, END)
        FechaAut_entry.delete(0, END)
        Desmotivo_entry.delete("1.0", END)
        Notas_entry.delete("1.0", END)
        dt_entry.delete(0, END)
        habilitar(0)
        my_tree.delete(selected)

    button_frame = LabelFrame(status_window, text="Comandos")
    button_frame.pack(fill="x", expand="yes", padx=20)

    dt_label = Label(button_frame, text="# DT Generado:")
    dt_label.grid(row=0, column=0, padx=10, pady=10)
    DT = tk.StringVar(status_window)
    dt_entry = tk.Entry(button_frame, textvariable=DT, width=10, validate="key",
                        validatecommand=(status_window.register(validation_dt), "%d", "%S", "%P"))
    dt_entry.grid(row=0, column=1, padx=10, pady=10)

    update_button = tk.Button(button_frame, text="Subir DT", command=update_record)
    update_button.grid(row=0, column=2, padx=10, pady=10)

    DT.trace_add("write", lambda name, index, mode: update_button_state())
    add_button = tk.Button(button_frame, text="Autorizar Regreso", command=autorizar, state=tk.DISABLED)
    add_button.grid(row=0, column=3, padx=10, pady=10)

    remove_all_button = tk.Button(button_frame, text="Rechazar Regreso", command=rechazar, state=tk.DISABLED)
    remove_all_button.grid(row=0, column=4, padx=10, pady=10)

    remove_one_button = tk.Button(button_frame, text="Remover selecionado", command=remove_one, state=tk.DISABLED)
    remove_one_button.grid(row=0, column=5, padx=10, pady=10)

    remove_many_button = tk.Button(button_frame, text="Remove varios", command=remove_many, state=tk.DISABLED)
    remove_many_button.grid(row=0, column=6, padx=10, pady=10)

    move_up_button = tk.Button(button_frame, text="Arriba", command=up, state=tk.DISABLED)
    move_up_button.grid(row=0, column=7, padx=10, pady=10)

    move_down_button = tk.Button(button_frame, text="Abajo", command=down, state=tk.DISABLED)
    move_down_button.grid(row=0, column=8, padx=10, pady=10)

    select_record_button = tk.Button(button_frame, text="Limpiar Registros", command=clear_entries, state=tk.NORMAL)
    select_record_button.grid(row=0, column=9, padx=10, pady=10)

    my_tree.bind("<ButtonRelease-1>", select_record)
    query_database(usuario)
    habilitar(0)
    status_window.protocol("WM_DELETE_WINDOW", close_window2)

