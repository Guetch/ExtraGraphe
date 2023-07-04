import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys
from tkinter import PhotoImage
from tkinter.filedialog import asksaveasfile, askopenfilename
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from PIL import ImageTk, Image
import webbrowser

# Fonction de création de page
def expo_page():
    global impo1_label_file
    global impo1_tv1
    expo_frame=tk.Frame(main_frame)
    lb=tk.Label(expo_frame, text="Expo Page\n\nPage:1", font=("Bold",30))
    lb.pack()
    expo_frame.pack(pady=18)

# création du contenu de la page
    impo1_frame = tk.LabelFrame(main_frame, text="Ouvrir un fichier Excel")
    impo1_frame.place(height=600, width=675)

    impo1_file=tk.LabelFrame(main_frame, text="Ouvrir")
    impo1_file.place(height=100, width=675, rely=0.85, relx=0)

    impo1_bouton=tk.Button(impo1_file, text="Choisir", command=impo1_importer)
    impo1_bouton.place(rely=0.65, relx=0.60)

    impo1_bouton1 = tk.Button(impo1_file, text="Ouvrir le fichier", command=impo1_ouvrir)
    impo1_bouton1.place(rely=0.65, relx=0.10)

    impo1_label_file=ttk.Label(impo1_file, text="Aucun fichier sélectionné")
    impo1_label_file.place(rely=0, relx=0)

    impo1_tv1=ttk.Treeview(impo1_frame)
    impo1_tv1.place(relheight=1, relwidth=1)
    # Création curseur
    impo1_treescrolly=tk.Scrollbar(impo1_frame,orient="vertical", command=impo1_tv1.yview)
    impo1_treescrollx = tk.Scrollbar(impo1_frame, orient="horizontal", command=impo1_tv1.xview)
    impo1_tv1.configure(xscrollcommand=impo1_treescrollx.set, yscrollcommand=impo1_treescrolly.set)
    impo1_treescrollx.pack(side="bottom", fill="x")
    impo1_treescrolly.pack(side="right", fill="y")

# Création des fonctions pour la page exporter
def impo1_importer():
    global impo1_label_file
    impo1_filename=filedialog.askopenfilename(initialdir="*/*", title="Ouvrir un fichier", filetype=(("Fichier Excel","*.xlsx"), ("Autre Fichier", "*.*")))
    impo1_label_file["text"] =impo1_filename
    return None

def impo1_ouvrir():
    global df
    impo1_file_path=impo1_label_file["text"]
    try:
        excel_filename=r"{}".format(impo1_file_path)
        df=pd.read_excel(excel_filename)
    except ValueError:
        tk.messagebox.showerror("Information", "Le fichier ouvert est invalide")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", "Fichier introuvable")
        return None

    clear_data()
    impo1_tv1["column"]=list(df.columns)
    impo1_tv1["show"]="headings"
    for column in impo1_tv1["column"]:
        impo1_tv1.heading(column, text=column)
    df.rows=df.to_numpy().tolist()
    for row in df.rows:
        impo1_tv1.insert("", tk.END, values=row)

def clear_data():
    global impo1_tv1
    impo1_tv1.delete(*impo1_tv1.get_children())

def donnee_page():
    global var1
    donnee_frame=tk.Frame(main_frame)
    lb=tk.Label(donnee_frame, text="Donnee Page\n\nPage:2", font=("Bold",30))
    lb.pack()
    donnee_frame.pack(pady=18)


    def add_two_input1(event):
        global Don1
        # Obtenir l'élément sélectionné
        donnee1_list = donnee1_listeCombo1.get()
        Don1=donnee1_listeCombo1.get()
        print(Don1)

    def add_two_input2(event):
        global Don2
        # Obtenir l'élément sélectionné
        donnee1_list = donnee1_listeCombo2.get()
        Don2=donnee1_listeCombo2.get()
        print(Don2)
    # création des cases à cochet pour des options
    var1 = tk.IntVar(value=0)
    def extract():
        global Don, extracted_data,  Don3
        x = df[Don1].values
        y = df[Don2].values
        if var1.get() == 0:
            try:
                A1 = donnee2_Entry_Min.get()
                A2 = donnee2_Entry_Max.get()
                A3 = donnee2_Entry_Pas.get()
                F = np.arange(float(A1), (float(A2)+float(A3)), float(A3))
            except ValueError:
                messagebox.showerror(title="erreur", message="Veuillez fournir 3 valeurs d'entrée valides")
        else:
            # Obtenir l'élément sélectionné
            donnee1_list = donnee3_listeCombo3.get()
            Don3 = donnee3_listeCombo3.get()
            print(Don3)
            print(var1.get())
            F = df[Don3].values
        x_to_extract = F
        y_extracted = np.interp(x_to_extract, x, y)
        Don = pd.DataFrame({"x_à_extraire": x_to_extract, "y_extrait": y_extracted})




    donnee1_frame = tk.LabelFrame(main_frame, text="Paramètre des données")
    donnee1_frame.place(height=1000, width=675)

    # Fénêtre du choix des colonnes de données d'entrée
    donnee1_file=tk.LabelFrame(main_frame, text="Choix des colonnes de coordonnées")
    donnee1_file.place(height=1000, width=675, rely=0.085, relx=0)

    donnee1_label_X=tk.Label(donnee1_file, text="X :")
    donnee1_label_X.grid(row=1, column=0, pady=10, padx=5, sticky="we")
    donnee1_label_Y=tk.Label(donnee1_file, text="Y :")
    donnee1_label_Y.grid(row=2, column=0, pady=10, padx=5, sticky="we")
# masque de combobox
    donnee1_listeCombo11 = ttk.Combobox(donnee1_file)
    donnee1_listeCombo11.grid(row=1, column=1, pady=10, padx=5, sticky="we")
    donnee1_listeCombo22 = ttk.Combobox(donnee1_file)
    donnee1_listeCombo22.grid(row=2, column=1, pady=10, padx=5, sticky="we")


    # Fénêtre des abscisse à extraire
    donnee2_file=tk.LabelFrame(main_frame, text="Enter les abscisses à extraire")
    donnee2_file.place(height=1000, width=675, rely=0.45, relx=0)





    donnee2_label_Min=tk.Label(donnee2_file, text="Min :")
    donnee2_label_Min.grid(row=1, column=0, pady=10, padx=5, sticky="we")
    donnee2_label_Max=tk.Label(donnee2_file, text="Max :")
    donnee2_label_Max.grid(row=2, column=0, pady=10, padx=5, sticky="we")
    donnee2_label_Pas=tk.Label(donnee2_file, text="p :")
    donnee2_label_Pas.grid(row=3, column=0, pady=10, padx=5, sticky="we")
    donnee2_label_Acces = tk.Label(donnee2_file, fg="red")
    donnee2_label_Acces.grid(row=4, column=0)

    donnee2_Entry_Min = tk.Entry(donnee2_file)
    donnee2_Entry_Min.grid(row=1, column=1)
    donnee2_Entry_Max = tk.Entry(donnee2_file)
    donnee2_Entry_Max.grid(row=2, column=1)
    donnee2_Entry_Pas = tk.Entry(donnee2_file)
    donnee2_Entry_Pas.grid(row=3, column=1)

    donnee2_bouton_ok = tk.Button(donnee2_file, text="Valider", command= lambda:extract())
    donnee2_bouton_ok.grid(row=4, column=1)

    donnee3_label_X=tk.Label(donnee2_file, text="Fichier Excel")
    donnee3_label_X.grid(row=6, column=0, pady=10, padx=5, sticky="we")
    donnee3_listeCombo13 = ttk.Combobox(donnee2_file)
    donnee3_listeCombo13.grid(row=6, column=1, pady=10, padx=5, sticky="we")

    donnee3_bouton_ok = tk.Button(donnee2_file, text="Valider", command= lambda:extract())
    donnee3_bouton_ok.grid(row=7, column=1)

# liste actif utilisé par le code
    donnee1_list = list(df[:0])
    donnee1_listeCombo1=ttk.Combobox(donnee1_file, values=donnee1_list, state="readonly")
    donnee1_listeCombo1.grid(row=1, column=1, pady=10, padx=5, sticky="we")
    donnee1_listeCombo1.bind("<<ComboboxSelected>>", add_two_input1)
    donnee1_listeCombo2=ttk.Combobox(donnee1_file, values=donnee1_list, state="readonly")
    donnee1_listeCombo2.grid(row=2, column=1, pady=10, padx=5, sticky="we")
    donnee1_listeCombo2.bind("<<ComboboxSelected>>", add_two_input2)

    donnee3_listeCombo3=ttk.Combobox(donnee2_file, values=donnee1_list, state="readonly")
    donnee3_listeCombo3.grid(row=6, column=1, pady=10, padx=5, sticky="we")
    donnee3_listeCombo3.set(donnee1_list[0])


    # create one check buttons
    donnee3_case = tk.Checkbutton(donnee2_file, text="Sélectionner dans la liste", variable=var1, onvalue=1, offvalue=0)
    # place the check buttons in the window
    donnee3_case.grid(row=6, column=3, pady=10, padx=5, sticky="we")


def resultat_page():
    global Don
    resultat_frame=tk.Frame(main_frame)
    lb=tk.Label(resultat_frame, text="Resultat Page\n\nPage:3", font=("Bold",30))
    lb.pack()
    resultat_frame.pack(pady=18)


    def affichage():
        global Don
        # Clear the treeview
        my_tree.delete(*my_tree.get_children())

        # get the headers
        my_tree["column"] = list(Don.columns)
        my_tree["show"] = "headings"

        # Show the headers
        for col in my_tree["column"]:
            my_tree.heading(col, text=col)

        # show data
        Don_rows = Don.to_numpy().tolist()
        for row in Don_rows:
            my_tree.insert("", "end", values=row)



    def plot():
        global x_to_extract, y_extracted
        ax.clear()
        X = Don["x_à_extraire"]
        Y = Don["y_extrait"]
        X1 = df[Don1].values
        Y1 = df[Don2].values
        ax.scatter(X, Y, label="Données extraites", color="red", s=10)
        ax.plot(X1, Y1, label="Données brutes")
        ax.set_xlabel("Axe x")
        ax.set_ylabel("Axe y")
        ax.set_title("Courbe comparative des données")
        ax.legend()
        canvas.draw()

    # Fenêtre du choix des colonnes de données extraites
    result1_file = tk.LabelFrame(main_frame, text="")
    result1_file.place(height=230, width=675, rely=0, relx=0)

    # Bouton pour afficher les données
    result1_bouton_ok = tk.Button(result1_file, text="Coordonnées", command=affichage)
    result1_bouton_ok.grid(row=0, column=0, sticky="w")

    # Création du Treeview
    my_tree = ttk.Treeview(result1_file)
    my_tree.grid(row=1, column=0, pady=5, padx=5, sticky="nsew")

    # Scrollbar verticale
    result1_treescrolly = tk.Scrollbar(result1_file, orient="vertical", command=my_tree.yview)
    result1_treescrolly.grid(row=1, column=1, sticky="ns")

    # Scrollbar horizontale
    result1_treescrollx = tk.Scrollbar(result1_file, orient="horizontal", command=my_tree.xview)
    result1_treescrollx.grid(row=2, column=0, sticky="ew")

    # Configuration de la scrollbar et de Treeview
    my_tree.configure(xscrollcommand=result1_treescrollx.set, yscrollcommand=result1_treescrolly.set)

    # Redimensionnement des lignes et colonnes
    result1_file.rowconfigure(1, weight=1)
    result1_file.columnconfigure(0, weight=1)

    # Fénêtre des abscisse à extraire
    result2_file=tk.LabelFrame(main_frame, text="")
    result2_file.place(height=525, width=675, rely=0.32, relx=0)
    result2_bouton_ok = tk.Button(result2_file, text="Graphe", command=plot)
    result2_bouton_ok.grid(row=0, column=0, sticky="w")

    # création de Canvas
    fig, ax = plt.subplots()
    canvas = FigureCanvasTkAgg(fig, master=result2_file)
    canvas.get_tk_widget().grid(row=1, column=0)
    canvas.get_tk_widget().configure(width=650, height=400)

    #curseur
    toolbar=NavigationToolbar2Tk(canvas, result2_file, pack_toolbar=False)
    toolbar.update()
    toolbar.grid(row=0, column=0)



# création de la fonction pour faire disparaitre les pages à la suite de louverture de l'autre
def delete_page():
    for frame in main_frame.winfo_children():
        frame.destroy()


# création de la fonction pour faire disparaitre les indicateurs
def hide_indicate():
    expo_indicate.config(bg="#c3c3c3")
    donnee_indicate.config(bg="#c3c3c3")
    resultat_indicate.config(bg="#c3c3c3")

# création de la fonction pour faire apparaitre les indicateurs
def indicate(lb, page):
    hide_indicate()
    lb.config(bg="#158aff")
    delete_page()
    page()



def enregistre():
    # ask the user to select a file to save the DataFrame
    save = filedialog.asksaveasfile(title="", filetypes=[("Fichier Excel", "*.xlsx"), ("Fichier texte", "*.txt"),
                                                         ("Autre fichier", "*.*")], defaultextension="*.xlsx", mode="w")
    if save is not None:
        try:
        # write the DataFrame to the Excel file
             Don.to_excel(save.name, index=False)

        # show a confirmation message
             messagebox.showinfo(title="Succès", message="Le fichier a été sauvegardé avec succès!")
        except:
        # show an error message
            messagebox.showerror(title="Erreur", message="Une erreur s'est produite lors de la sauvegarde du fichier.")
        finally:
        # close the file
            save.close()
    else:
    # show an error message
        messagebox.showerror(title="Erreur", message="Aucun fichier n'a été sélectionné.")
# Création de la fenêtre du message à propos
def show_a_propos():
    messagebox.showinfo("A propos", "ExtraGraphe: Logiciel de ré-échantillonnage par extraction de coordonnées sur une courbe ou une droite 2D\nVersion : 1.0\nAuteur : GUETCHAMEGNI Elvis Le Doux")

def show_Tutoriel():
    messagebox.showinfo("Tutoriel", "Cliquer sur le lien : https://www.elittedeco.fr/contact/ ")

def open_download_link():
    webbrowser.open("http://extragraphe.local/tutoriel/")
def open_download_link1():
    webbrowser.open("https://www.paypal.com/donate?token=ALDAih_N9laAJEhIYDywRksaN1z3ubNFyZlrzx1IhgXvYsZawqIAcf7xdKsJ7MNILqzW5m04Iq2YDx3J")


fen=tk.Tk()
fen.title("ExtraGraphe")
# Changer le logo (https://icoconvert.com/, pour convertir l'image png ou autre)
fen.iconbitmap(r"C:/Users/Utilisateur/PycharmProjects/Guetch_CD/logo_extragraphe.ico")
fen.geometry("800x1500")

# Création du menu
# créer la barre de menu
menu_h=tk.Menu(fen)

# créer la cascade
menu_fichier=tk.Menu(menu_h, tearoff=0)
menu_aide=tk.Menu(menu_h, tearoff=0)

menu_h.add_cascade(label="Fichier", menu=menu_fichier)
menu_h.add_cascade(label="Aide", menu=menu_aide)

# créer le contenu
menu_fichier.add_command(label="Ouvrir", command=expo_page)
menu_fichier.add_command(label="Exporter", command=enregistre)
menu_fichier.add_command(label="Quitter", command=sys.exit)
menu_aide.add_command(label="A propos", command=show_a_propos)
menu_aide.add_command(label="Tutoriel", command=open_download_link)
menu_aide.add_command(label="Faire un don (PayPal)", command=open_download_link1)

# placement



fen.config(menu=menu_h)


# Création de la mage
options_frame=tk.Frame(fen, bg="#c3c3c3")


options_frame.pack(side=tk.LEFT)
options_frame.pack_propagate(False)
options_frame.config(width=100, height=1400)

# Charger l'image
image = Image.open("C:/Users/Utilisateur/PycharmProjects/extact/image.png")
bg_image = ImageTk.PhotoImage(image)



main_frame=tk.Frame(fen, highlightbackground="black", highlightthickness=2)
main_frame.pack(side=tk.LEFT)
main_frame.pack_propagate(False)
main_frame.config(height=1500,width=700)

background_label = tk.Label(main_frame, image=bg_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

# Création de bouton
expo_btn=tk.Button(options_frame, text="Import", font=("Bold", 15), fg="#158aff", bd=0, bg="#c3c3c3", command=lambda:indicate(expo_indicate, expo_page))
expo_btn.place(x=10, y=100)

donnee_btn=tk.Button(options_frame, text="Donnée", font=("Bold", 15), fg="#158aff", bd=0, bg="#c3c3c3", command=lambda:indicate(donnee_indicate, donnee_page))
donnee_btn.place(x=10, y=200)

resultat_btn=tk.Button(options_frame, text="Résultat", font=("Bold", 15), fg="#158aff", bd=0, bg="#c3c3c3", command=lambda:indicate(resultat_indicate, resultat_page))
resultat_btn.place(x=10, y=300)

# création de indicateur
expo_indicate=tk.Label(options_frame,text="",bg="#c3c3c3")
expo_indicate.place(x=3, y=100, width=6, height=40)

donnee_indicate=tk.Label(options_frame,text="",bg="#c3c3c3")
donnee_indicate.place(x=3, y=200, width=6, height=40)

resultat_indicate=tk.Label(options_frame,text="",bg="#c3c3c3")
resultat_indicate.place(x=3, y=300, width=6, height=40)



fen.mainloop()


