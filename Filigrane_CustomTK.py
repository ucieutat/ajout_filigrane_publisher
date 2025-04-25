import os
import logging
import customtkinter as ctk
import win32com.client
from tkinter import filedialog, Tk

# Configurer les logs
logging.basicConfig(
    filename="programme.log",
    filemode="w",
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s"
)


# 💠 Interface principale CustomTkinter
class FiligraneAppGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Filigrane Publisher")
        self.geometry("400x250")
        self.resizable(False, False)

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.user_input = None

        self.label = ctk.CTkLabel(self, text="Entrez le nom du filigrane :", font=("Arial", 14))
        self.label.pack(pady=20)

        self.entry = ctk.CTkEntry(self, placeholder_text="Nom du filigrane")
        self.entry.pack(pady=10)

        self.button = ctk.CTkButton(self, text="Valider", command=self.on_submit)
        self.button.pack(pady=10)

        self.theme_button = ctk.CTkButton(self, text="🌙 / ☀️", command=self.toggle_theme, width=50, height=25, anchor="center")
        self.theme_button.place(x=10, y=10)

        self.mode = ctk.get_appearance_mode()

    def toggle_theme(self):
        # Basculer le mode actuel
        if self.mode == "light":
            ctk.set_appearance_mode("dark")
            self.mode = "dark"  # Mettre à jour l'état du mode
        else:
            ctk.set_appearance_mode("light")
            self.mode = "light"

    def on_submit(self):
        self.user_input = self.entry.get()
        self.after(100, self._exit_safe)

    def _exit_safe(self):
        self.quit()
        self.destroy()

    @staticmethod
    def afficher_message_fin():
        msgbox = ctk.CTk()

        def close():
            msgbox.after(100, lambda: [msgbox.quit(), msgbox.destroy()])

        msgbox.title("Terminé")
        msgbox.geometry("300x150")
        msgbox.resizable(False, False)

        label = ctk.CTkLabel(msgbox, text="Travail terminé\nMerci à tous !", font=("Arial", 16))
        label.pack(pady=30)

        bouton = ctk.CTkButton(msgbox, text="OK", command=close)
        bouton.pack(pady=10)

        msgbox.mainloop()


# 📂 Sélection de fichiers
def demander_chemin_fichier(message, type_fichier):
    try:
        logging.info("Ouverture de la boîte de dialogue pour sélectionner des fichiers.")
        root = Tk()
        root.withdraw()
        fichiers = filedialog.askopenfilenames(title=message, filetypes=[('Fichiers', type_fichier)])
        root.destroy()
        fichiers_normaux = [os.path.normpath(f.replace('%20', ' ')) for f in fichiers]
        logging.info(f"Fichiers sélectionnés : {fichiers_normaux}")
        return fichiers_normaux
    except Exception as e:
        logging.error(f"Erreur lors de la sélection du fichier : {e}")
        return []


# 📁 Création du dossier destination
def creer_dossier(nom_utilisateur):
    dossier = os.path.normpath(os.path.join(os.getcwd(), f"Situations_problemes_{nom_utilisateur}"))
    if not os.path.exists(dossier):
        os.makedirs(dossier)
        logging.info(f"Dossier créé : {dossier}")
    else:
        logging.info(f"Dossier déjà existant : {dossier}")
    return dossier


# 🖨️ Traitement des fichiers Publisher
def ajouter_filigrane_et_exporter(app, fichiers, destination, filigrane):
    PbFixedFormatType_pbFixedFormatTypePDF = 2
    PbFixedFormatIntent_pbIntentStandard = 2

    for fichier in fichiers:
        try:
            logging.info(f"Ouverture du fichier Publisher : {fichier}")
            app.Open(fichier)
            for page in range(2):  # Suppose 2 pages
                textbox = app.ActiveDocument.Pages(page + 1).Shapes.AddTextbox(
                    Orientation=2, Left=21, Top=330, Width=10, Height=300
                )
                textbox.TextFrame.TextRange.Text = f"Copie personnelle de : {filigrane}"
                logging.info(f"Filigrane ajouté à la page {page + 1} de {fichier}")

            pdf_path = os.path.join(destination, os.path.basename(fichier)[:-4] + ".pdf")
            app.ActiveDocument.ExportAsFixedFormat(
                PbFixedFormatType_pbFixedFormatTypePDF, pdf_path,
                PbFixedFormatIntent_pbIntentStandard, True
            )
            logging.info(f"Exporté en PDF : {pdf_path}")
            app.ActiveDocument.Close()
        except Exception as e:
            logging.error(f"Erreur sur le fichier {fichier} : {e}")


# ▶️ Fonction principale
def main():
    try:
        logging.info("Démarrage du programme.")
        app_pub = win32com.client.Dispatch("Publisher.Application")
        app_pub.ActiveWindow.Visible = False

        # Interface CustomTkinter
        gui = FiligraneAppGUI()
        gui.mainloop()
        texte_utilisateur = gui.user_input

        if not texte_utilisateur:
            raise ValueError("Aucune entrée utilisateur détectée.")

        destination = creer_dossier(texte_utilisateur)
        fichiers = demander_chemin_fichier("Sélectionnez les fichiers Publisher", "*.pub")

        if not fichiers:
            raise ValueError("Aucun fichier sélectionné.")

        ajouter_filigrane_et_exporter(app_pub, fichiers, destination, texte_utilisateur)

        FiligraneAppGUI.afficher_message_fin()
        logging.info("Travail terminé avec succès.")

    except Exception as e:
        logging.critical(f"Erreur critique : {e}")

    finally:
        if 'app_pub' in locals():
            app_pub.Quit()
            logging.info("Application Publisher fermée.")


if __name__ == "__main__":
    main()
