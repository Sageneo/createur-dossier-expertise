import os
import shutil
import re
import email
import datetime
from email.parser import BytesParser
from email.policy import default
import tkinter as tk
from tkinter import filedialog, messagebox

def extraire_informations_mail(chemin_fichier_eml):
    """Extrait les informations clés d'un fichier .eml"""
    
    # Ouvrir et parser le fichier .eml
    with open(chemin_fichier_eml, 'rb') as fp:
        msg = BytesParser(policy=default).parse(fp)
    
    # Extraire la date
    date_str = msg.get('Date')
    date_envoi = None
    
    if date_str:
        # Essayer de parser la date du format standard d'email
        try:
            # Convertir la date en objet datetime
            date_tuple = email.utils.parsedate_tz(date_str)
            if date_tuple:
                date_envoi = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
        except:
            # Si échec, utiliser la date actuelle
            date_envoi = datetime.datetime.now()
    else:
        # Si pas de date, utiliser la date actuelle
        date_envoi = datetime.datetime.now()
    
    # Extraire le sujet et le contenu
    sujet = msg.get('Subject', '')
    
    # Récupérer le contenu du mail
    contenu = ""
    
    # Récupérer le corps du message
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            
            # Ignorer les pièces jointes
            if "attachment" not in content_disposition:
                if content_type == "text/plain":
                    contenu += part.get_payload(decode=True).decode('utf-8', errors='ignore')
                elif content_type == "text/html":
                    # Si on a du HTML et pas encore de texte, on le garde
                    if not contenu:
                        contenu += part.get_payload(decode=True).decode('utf-8', errors='ignore')
    else:
        contenu = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
    
    # Extraire la ville
    match_ville = re.search(r'Ville:\s*(\w+)', contenu)
    if not match_ville:
        match_ville = re.search(r'Code postal:\s*(\d+)\s+Ville:\s*(\w+)', contenu)
        if match_ville:
            ville = match_ville.group(2)
        else:
            ville = "Inconnue"
    else:
        ville = match_ville.group(1)
    
    # Extraire le nom de contact
    match_nom = re.search(r'Contact:\s*([\w\s-]+)', contenu)
    if not match_nom:
        match_nom = re.search(r'Nom/Entreprise:\s*([\w\s-]+)', contenu)
    
    nom_contact = match_nom.group(1).strip() if match_nom else "Inconnu"
    
    # Nettoyer le nom de contact (enlever les caractères spéciaux)
    nom_contact = re.sub(r'[\\/*?:"<>|]', '', nom_contact)
    
    # Extraire les pièces jointes
    pieces_jointes = []
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            
            filename = part.get_filename()
            if filename:
                pieces_jointes.append({
                    'filename': filename,
                    'content': part.get_payload(decode=True),
                    'content_type': part.get_content_type()
                })
    
    # Si on ne trouve pas de pièces jointes avec la méthode ci-dessus,
    # essayer de les extraire du texte
    if not pieces_jointes:
        pieces_jointes_noms = re.findall(r'Document \d+: ([\w.-]+\.(?:png|jpg|pdf|doc|docx|xls|xlsx))', contenu)
        pieces_jointes = [{'filename': nom, 'content': None, 'content_type': None} for nom in pieces_jointes_noms]
    
    return {
        'date': date_envoi,
        'ville': ville,
        'nom_contact': nom_contact,
        'pieces_jointes': pieces_jointes,
        'contenu': contenu
    }

def est_image_ou_plan(nom_fichier, type_contenu):
    """Détermine si un fichier est une image ou un plan"""
    extensions_image = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff']
    extensions_plan = ['.pdf', '.dwg', '.dxf']
    
    # Vérifier l'extension
    _, ext = os.path.splitext(nom_fichier.lower())
    
    # Vérifier le type de contenu
    est_image = (ext in extensions_image) or (type_contenu and 'image' in type_contenu)
    est_plan = (ext in extensions_plan) or ('plan' in nom_fichier.lower())
    
    return est_image or est_plan

def sauvegarder_pieces_jointes(pieces_jointes, dossier_echange_mails, dossier_visite):
    """Sauvegarde les pièces jointes dans les dossiers appropriés"""
    
    pieces_jointes_sauvegardees = []
    
    for piece in pieces_jointes:
        filename = piece['filename']
        content = piece['content']
        content_type = piece['content_type']
        
        if not filename:
            continue
        
        # Nettoyer le nom de fichier
        filename = re.sub(r'=\?.*\?=', '', filename).strip()
        # Supprimer les caractères non autorisés dans les noms de fichiers Windows
        filename = re.sub(r'[\\/*?:"<>|]', '', filename)
        
        # Déterminer le dossier de destination
        if est_image_ou_plan(filename, content_type):
            dossier_destination = os.path.join(dossier_visite, "Photos")
        else:
            dossier_destination = os.path.join(dossier_echange_mails, "Pièces jointes")
        
        # S'assurer que le dossier existe
        if not os.path.exists(dossier_destination):
            os.makedirs(dossier_destination)
        
        # Sauvegarder la pièce jointe si on a son contenu
        if content:
            chemin_fichier = os.path.join(dossier_destination, filename)
            with open(chemin_fichier, 'wb') as f:
                f.write(content)
            
            pieces_jointes_sauvegardees.append({
                'nom': filename,
                'chemin': chemin_fichier
            })
    
    return pieces_jointes_sauvegardees

def nettoyer_nom_fichier(nom):
    """Nettoie un nom de fichier pour éviter les erreurs"""
    # Remplacer les retours à la ligne
    nom = nom.replace('\n', '')
    # Remplacer les caractères non autorisés dans les noms de fichiers Windows
    nom = re.sub(r'[\\/*?:"<>|]', '', nom)
    return nom

def renommer_fichiers_et_dossiers(dossier_racine, ancien_motif, nouveau_motif, nouveau_motif_excel=None):
    """Renomme tous les fichiers et dossiers contenant un motif spécifique"""
    
    # Nettoyer les motifs
    ancien_motif = nettoyer_nom_fichier(ancien_motif)
    nouveau_motif = nettoyer_nom_fichier(nouveau_motif)
    if nouveau_motif_excel:
        nouveau_motif_excel = nettoyer_nom_fichier(nouveau_motif_excel)
    
    # Liste pour stocker tous les chemins à renommer (pour éviter les problèmes pendant l'itération)
    elements_a_renommer = []
    
    # Parcourir tous les fichiers et dossiers
    for dossier_courant, sous_dossiers, fichiers in os.walk(dossier_racine, topdown=False):
        # Ajouter les fichiers à renommer
        for fichier in fichiers:
            if ancien_motif in fichier:
                chemin_complet = os.path.join(dossier_courant, fichier)
                
                # Vérifier si c'est un fichier Excel dans le dossier "5 - Rapport AVT"
                if "5 - Rapport AVT" in dossier_courant and fichier.endswith(".xlsx") and nouveau_motif_excel:
                    # Utiliser le format spécial pour le fichier Excel
                    nouveau_nom = fichier.replace(ancien_motif, nouveau_motif_excel)
                else:
                    # Utiliser le format standard pour les autres fichiers
                    nouveau_nom = fichier.replace(ancien_motif, nouveau_motif)
                
                elements_a_renommer.append((chemin_complet, os.path.join(dossier_courant, nouveau_nom)))
        
        # Ajouter les dossiers à renommer
        for sous_dossier in sous_dossiers:
            if ancien_motif in sous_dossier:
                chemin_complet = os.path.join(dossier_courant, sous_dossier)
                nouveau_nom = sous_dossier.replace(ancien_motif, nouveau_motif)
                elements_a_renommer.append((chemin_complet, os.path.join(dossier_courant, nouveau_nom)))
    
    # Renommer tous les éléments (en commençant par les plus profonds grâce à topdown=False)
    for ancien_chemin, nouveau_chemin in elements_a_renommer:
        try:
            os.rename(ancien_chemin, nouveau_chemin)
            print(f"Renommé: {os.path.basename(ancien_chemin)} -> {os.path.basename(nouveau_chemin)}")
        except Exception as e:
            print(f"Erreur lors du renommage de {ancien_chemin}: {e}")

def creer_dossier_expertise(infos, dossier_modele, dossier_destination, chemin_fichier_eml):
    """Crée un nouveau dossier d'expertise basé sur le modèle"""
    
    # Format de la date: YYMMDD
    date_format = infos['date'].strftime('%y%m%d')
    
    # Motif à rechercher dans les noms de fichiers/dossiers
    ancien_motif = "Num - Ville - Sociétée"
    
    # Création du nom de dossier selon le format demandé
    nouveau_motif = f"00 - {date_format} - {infos['ville']} - {infos['nom_contact']}"
    
    # Format spécial pour le fichier Excel (sans les "00 - " au début)
    nouveau_motif_excel = f"{date_format} - {infos['ville']} - {infos['nom_contact']}"
    
    # Supprimer le mot "Email" à la fin des noms s'il est présent
    nouveau_motif = nouveau_motif.replace("Email", "").strip()
    nouveau_motif_excel = nouveau_motif_excel.replace("Email", "").strip()
    
    # Nettoyer les noms pour éviter les erreurs
    nouveau_motif = nettoyer_nom_fichier(nouveau_motif)
    nouveau_motif_excel = nettoyer_nom_fichier(nouveau_motif_excel)
    
    # Nom du dossier modèle (juste le nom, pas le chemin complet)
    nom_dossier_modele = os.path.basename(dossier_modele)
    
    # Chemin complet du nouveau dossier
    chemin_nouveau_dossier = os.path.join(dossier_destination, nouveau_motif)
    
    # Vérifier si le dossier existe déjà
    if os.path.exists(chemin_nouveau_dossier):
        print(f"Le dossier {nouveau_motif} existe déjà.")
        return chemin_nouveau_dossier
    
    # Copier le dossier modèle
    shutil.copytree(dossier_modele, chemin_nouveau_dossier)
    print(f"Dossier créé: {nouveau_motif}")
    
    # Renommer tous les fichiers et dossiers contenant l'ancien motif
    renommer_fichiers_et_dossiers(chemin_nouveau_dossier, ancien_motif, nouveau_motif, nouveau_motif_excel)
    
    # Créer les dossiers nécessaires s'ils n'existent pas
    dossier_echange_mails = os.path.join(chemin_nouveau_dossier, "1 - Echange de mails")
    dossier_visite = os.path.join(chemin_nouveau_dossier, "3 - Visite")
    
    if not os.path.exists(dossier_echange_mails):
        os.makedirs(dossier_echange_mails)
    
    if not os.path.exists(dossier_visite):
        os.makedirs(dossier_visite)
    
    # Créer un dossier pour les pièces jointes dans le dossier d'échange de mails
    dossier_pj = os.path.join(dossier_echange_mails, "Pièces jointes")
    if not os.path.exists(dossier_pj):
        os.makedirs(dossier_pj)
    
    # Créer un dossier pour les photos dans le dossier de visite
    dossier_photos = os.path.join(dossier_visite, "Photos")
    if not os.path.exists(dossier_photos):
        os.makedirs(dossier_photos)
    
    # Extraire et sauvegarder les pièces jointes
    pieces_jointes = sauvegarder_pieces_jointes(
        infos['pieces_jointes'], 
        dossier_echange_mails, 
        dossier_visite
    )
    
    # Sauvegarder une copie du mail dans le dossier d'échange de mails
    nom_fichier_eml = os.path.basename(chemin_fichier_eml)
    shutil.copy2(chemin_fichier_eml, os.path.join(dossier_echange_mails, nom_fichier_eml))
    
    # Créer un fichier texte avec le contenu du mail
    with open(os.path.join(dossier_echange_mails, "contenu_mail.txt"), 'w', encoding='utf-8') as f:
        f.write(infos['contenu'])
    
    return chemin_nouveau_dossier

def interface_graphique():
    """Crée une interface graphique simple pour sélectionner les fichiers et dossiers"""
    
    def selectionner_fichier_eml():
        fichier = filedialog.askopenfilename(
            title="Sélectionner le fichier email (.eml)",
            filetypes=[("Fichiers Email", "*.eml"), ("Tous les fichiers", "*.*")]
        )
        if fichier:
            entry_fichier_eml.delete(0, tk.END)
            entry_fichier_eml.insert(0, fichier)
    
    def selectionner_dossier_modele():
        dossier = filedialog.askdirectory(title="Sélectionner le dossier modèle à copier")
        if dossier:
            entry_dossier_modele.delete(0, tk.END)
            entry_dossier_modele.insert(0, dossier)
    
    def selectionner_dossier_destination():
        dossier = filedialog.askdirectory(title="Sélectionner le dossier de destination")
        if dossier:
            entry_dossier_destination.delete(0, tk.END)
            entry_dossier_destination.insert(0, dossier)
    
    def traiter_email():
        fichier_eml = entry_fichier_eml.get()
        dossier_modele = entry_dossier_modele.get()
        dossier_destination = entry_dossier_destination.get()
        
        if not fichier_eml or not os.path.isfile(fichier_eml):
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier email valide.")
            return
        
        if not dossier_modele or not os.path.isdir(dossier_modele):
            messagebox.showerror("Erreur", "Veuillez sélectionner un dossier modèle valide.")
            return
        
        if not dossier_destination or not os.path.isdir(dossier_destination):
            messagebox.showerror("Erreur", "Veuillez sélectionner un dossier de destination valide.")
            return
        
        try:
            # Extraire les informations du mail
            infos = extraire_informations_mail(fichier_eml)
            
            # Supprimer le mot "Email" du nom de contact s'il est présent
            infos['nom_contact'] = infos['nom_contact'].replace("Email", "").strip()
            
            # Compter les images/plans pour l'affichage
            nb_images = sum(1 for pj in infos['pieces_jointes'] 
                           if est_image_ou_plan(pj['filename'], pj['content_type']))
            nb_autres = len(infos['pieces_jointes']) - nb_images
            
            # Afficher les informations extraites pour vérification
            confirmation = messagebox.askyesno(
                "Confirmation",
                f"Informations extraites du mail:\n\n"
                f"Date: {infos['date'].strftime('%d/%m/%Y')}\n"
                f"Ville: {infos['ville']}\n"
                f"Contact: {infos['nom_contact']}\n"
                f"Images/Plans: {nb_images} (seront placés dans '3 - Visite/Photos')\n"
                f"Autres pièces jointes: {nb_autres} (seront placées dans '1 - Echange de mails/Pièces jointes')\n\n"
                f"Ces informations sont-elles correctes?"
            )
            
            if confirmation:
                # Créer le dossier d'expertise
                nouveau_dossier = creer_dossier_expertise(
                    infos, 
                    dossier_modele, 
                    dossier_destination,
                    fichier_eml
                )
                
                messagebox.showinfo(
                    "Succès", 
                    f"Dossier d'expertise créé avec succès:\n{nouveau_dossier}\n\n"
                    f"• Le mail a été placé dans '1 - Echange de mails'\n"
                    f"• Les images et plans ont été placés dans '3 - Visite/Photos'\n"
                    f"• Les autres pièces jointes sont dans '1 - Echange de mails/Pièces jointes'\n"
                    f"• Le fichier Excel dans '5 - Rapport AVT' a été renommé avec le format:\n"
                    f"  {infos['date'].strftime('%y%m%d')} - {infos['ville']} - {infos['nom_contact']}.xlsx"
                )
                
                # Demander si l'utilisateur veut ouvrir le dossier
                ouvrir_dossier = messagebox.askyesno(
                    "Ouvrir le dossier",
                    f"Voulez-vous ouvrir le dossier créé?"
                )
                
                if ouvrir_dossier:
                    # Ouvrir le dossier dans l'explorateur de fichiers
                    os.startfile(nouveau_dossier) if os.name == 'nt' else os.system(f'xdg-open "{nouveau_dossier}"')
                
                # Réinitialiser les champs
                entry_fichier_eml.delete(0, tk.END)
        
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite:\n{str(e)}")
            # Afficher plus de détails dans la console pour le débogage
            import traceback
            traceback.print_exc()
    
    # Créer la fenêtre principale
    fenetre = tk.Tk()
    fenetre.title("Création de dossier d'expertise")
    fenetre.geometry("600x380")
    
    # Créer les widgets
    frame = tk.Frame(fenetre, padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    # Fichier email
    tk.Label(frame, text="Fichier email (.eml):").grid(row=0, column=0, sticky="w", pady=5)
    entry_fichier_eml = tk.Entry(frame, width=50)
    entry_fichier_eml.grid(row=0, column=1, padx=5, pady=5)
    tk.Button(frame, text="Parcourir...", command=selectionner_fichier_eml).grid(row=0, column=2, padx=5, pady=5)
    
    # Dossier modèle
    tk.Label(frame, text="Dossier modèle:").grid(row=1, column=0, sticky="w", pady=5)
    entry_dossier_modele = tk.Entry(frame, width=50)
    entry_dossier_modele.grid(row=1, column=1, padx=5, pady=5)
    tk.Button(frame, text="Parcourir...", command=selectionner_dossier_modele).grid(row=1, column=2, padx=5, pady=5)
    
    # Dossier destination
    tk.Label(frame, text="Dossier destination:").grid(row=2, column=0, sticky="w", pady=5)
    entry_dossier_destination = tk.Entry(frame, width=50)
    entry_dossier_destination.grid(row=2, column=1, padx=5, pady=5)
    tk.Button(frame, text="Parcourir...", command=selectionner_dossier_destination).grid(row=2, column=2, padx=5, pady=5)
    
    # Instructions
    instructions = (
        "Instructions:\n\n"
        "1. Sélectionnez le fichier email (.eml) contenant les informations d'expertise\n"
        "2. Sélectionnez le dossier modèle à copier\n"
        "3. Sélectionnez le dossier où créer la nouvelle structure\n"
        "4. Cliquez sur 'Créer le dossier d'expertise'\n\n"
        "Le programme placera automatiquement:\n"
        "• Le mail dans '1 - Echange de mails'\n"
        "• Les images et plans dans '3 - Visite/Photos'\n"
        "• Les autres pièces jointes dans '1 - Echange de mails/Pièces jointes'\n"
        "• Renommera le fichier Excel dans '5 - Rapport AVT' avec le format:\n"
        "  YYMMDD - Ville - Nom.xlsx"
    )
    tk.Label(frame, text=instructions, justify=tk.LEFT, anchor="w").grid(row=3, column=0, columnspan=3, sticky="w", pady=10)
    
    # Bouton de traitement
    tk.Button(
        frame, 
        text="Créer le dossier d'expertise", 
        command=traiter_email,
        bg="#4CAF50",
        fg="white",
        padx=10,
        pady=5,
        font=("Arial", 10, "bold")
    ).grid(row=4, column=0, columnspan=3, pady=10)
    
    # Lancer l'interface
    fenetre.mainloop()

def main():
    """Fonction principale"""
    interface_graphique()

if __name__ == "__main__":
    main()
