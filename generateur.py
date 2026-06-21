import csv
import os
import urllib.parse

output_dir = "signatures_sorties"
os.makedirs(output_dir, exist_ok=True)

# Dictionnaire complet de tes logos
LOGOS_ENTREPRISES = {
    "michon 51": "https://www.trema-multitechnique.com/wp-content/uploads/2026/06/mail_MICHON-REIMS-scaled.png",
    "michon reims": "https://www.trema-multitechnique.com/wp-content/uploads/2026/06/mail_MICHON-REIMS-scaled.png",
    "michon idf": "https://www.trema-multitechnique.com/wp-content/uploads/2026/06/mail_MICHON-IDF-scaled.png",
    "michon 77": "https://www.trema-multitechnique.com/wp-content/uploads/2026/06/mail_MICHON-IDF-scaled.png",
    "duet conception": "https://www.trema-multitechnique.com/wp-content/uploads/2026/06/DUET-CONCEPTION-scaled.png",
    "duet tce": "https://www.trema-multitechnique.com/wp-content/uploads/2026/06/DUET-TCE-scaled.png",
    "trema": "https://www.trema-multitechnique.com/wp-content/uploads/2026/06/mail_TREMA-MULTITECHNIQUE-scaled.png",
    "ap2i": "https://www.trema-multitechnique.com/wp-content/uploads/2026/06/mail_AP2I-20-scaled.png",
    "etel": "https://www.trema-multitechnique.com/wp-content/uploads/2026/06/mail_ETEL-18-scaled.png",
    "michon": "https://www.trema-multitechnique.com/wp-content/uploads/2026/06/mail_MICHON-09-scaled.png",
    "secabat": "https://www.trema-multitechnique.com/wp-content/uploads/2026/06/mail_SECABAT-21-scaled.png",
    "guerineau": "https://www.trema-multitechnique.com/wp-content/uploads/2026/06/GUERINEAU_1-scaled.png",
    "deoust": "https://www.trema-multitechnique.com/wp-content/uploads/2026/06/mail_DEOUST-17-scaled.png",
}

LOGO_PAR_DEFAUT = LOGOS_ENTREPRISES["trema"]

# Dictionnaire complet de tes couleurs
COULEURS_ENTREPRISES = {
    "trema": "#f7a420",
    "etel": "#2aabe2",
    "deoust": "#d3222a",
    "secabat": "#ca6e26",
    "michon": "#008342",
    "duet": "#6b1416",
    "ap2i": "#0067b9",
    "guerineau": "#253759"
}

# Modèle HTML original préservé
html_template = """
<table cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width: 600px; font-family: Arial, sans-serif;">
  <tr>
    <td valign="middle" width="220" style="width: 220px; background-color: #ffffff; padding: 0; margin: 0;">
      <a href="https://www.trema-multitechnique.com" style="text-decoration: none !important; display: block; border: 0;">
        <img src="{logo}" alt="Logo Entreprise" width="220" style="display: block; width: 220px; border: 0; margin: 0; padding: 0;">
      </a>
    </td>
    
    <td valign="middle" bgcolor="{couleur}" style="background-color: {couleur}; color: #ffffff; padding: 15px 20px;">
      
      <div style="font-size: 18px; font-weight: bold; color: #ffffff; line-height: 1.2; margin: 0; padding: 0;">{prenom} <span style="text-transform: uppercase;">{nom}</span></div>      
      <div style="font-size: 14px; color: #ffffff; line-height: 1.2; margin: 2px 0 0 0; padding: 0;">{poste}</div>
      
      <p style="font-size: 4px; line-height: 4px; margin: 0; padding: 0;"> </p>
      
      <div style="font-size: 13px; line-height: 1.6; color: #ffffff; margin: 0; padding: 0;">
        {bloc_contact}
      </div>
      
      <div style="font-size: 11px; line-height: 11px; color: #ffffff; letter-spacing: -1px; margin: 0; padding: 0; white-space: nowrap; overflow: hidden; opacity: 0.9;">
        _______________________________________________________________
      </div>
      
      <p style="font-size: 6px; line-height: 6px; margin: 0; padding: 0;"> </p>
      
      <div style="font-size: 13px; line-height: 1.2; color: #ffffff; margin: 0; padding: 0;">
        <a href="https://www.trema-multitechnique.com" style="font-size: 15px; color: #ffffff !important; text-decoration: underline !important; font-weight: bold; display: inline-block; vertical-align: middle;">»</a><span style="vertical-align: middle;">  </span><span style="font-weight: bold; text-transform: uppercase; vertical-align: middle; color: #ffffff !important; text-decoration: none !important;">www<span style="font-family: Arial, sans-serif; font-size: 1px; line-height: 1px;"> </span>.trema-multitechnique.com</span>
      </div>
      
    </td>
  </tr>
</table>
"""

nom_fichier_csv = 'employes.csv'

def decoder_excel_texte(texte):
    if not texte: 
        return ""
    corrections = {
        "Ã©": "é", "Ã¨": "è", "Ã ": "à", "Ã§": "ç", "Ã¯": "ï", "Ã®": "î", "Ã»": "û", 
        "Ã”": "Ô", "â€“": "–", "\x82": "é", "\x90": "é", "\x8a": "è", "\x85": "à", 
        "\x87": "ç", "\x8b": "ï", "\x96": "û"
    }
    for broken, fixed in corrections.items():
        texte = texte.replace(broken, fixed)
    return texte.strip()

try:
    with open(nom_fichier_csv, mode='r', encoding='latin1', errors='replace') as file:
        reader = csv.DictReader(file, delimiter=';')
        
        # Nettoyage strict des espaces invisibles autour des noms de colonnes du CSV
        reader.fieldnames = [f.strip() if f else "" for f in reader.fieldnames]
        
        compteur = 0
        for row in reader:
            # Extraction directe et ciblée sur tes colonnes corrigées
            prenom = decoder_excel_texte(row.get('Prénom') or row.get('prénom') or '')
            nom = decoder_excel_texte(row.get('Nom') or row.get('nom') or '')
            poste = decoder_excel_texte(row.get('Poste') or row.get('poste') or '')
            telephone = decoder_excel_texte(row.get('numéro de téléphone') or row.get('Numéro de téléphone') or '').strip()
            email = decoder_excel_texte(row.get('Mail') or row.get('mail') or '').strip()
            adresse = decoder_excel_texte(row.get('Adresse') or row.get('adresse') or '')
            entreprise_brute = decoder_excel_texte(row.get('Entreprise') or row.get('entreprise') or '').lower().strip()
            
            # Sauter la ligne si elle est totalement vide
            if not nom and not prenom and not poste:
                continue
            
            # Nettoyage pour les noms de fichiers HTML
            entreprise_nettoyee = entreprise_brute.replace(" ", "_").replace("'", "")
            if not entreprise_nettoyee: 
                entreprise_nettoyee = "entreprise"

            prenom_clean = prenom.lower().strip().replace(" ", "_").replace("'", "") if prenom else "inconnu"
            nom_clean = nom.lower().strip().replace(" ", "_").replace("'", "") if nom else "inconnu"
            
            # Formatage pour l'affichage final
            prenom_affichage = prenom.capitalize() if prenom else " "
            nom_affichage = nom.upper() if nom else ""
            
            lignes_contact = []
            
            # 1. EMAIL
            if email and "@" in email:
                gauche, droite = email.split("@", 1)
                email_truque = f'{gauche}@<span style="font-family: Arial, sans-serif; font-size: 1px; line-height: 1px;"> </span>{droite}'
                lignes_contact.append(
                    f'<a href="mailto:{email}" style="font-size: 14px; color: #ffffff !important; text-decoration: underline !important; display: inline-block; vertical-align: middle;">✉</a>'
                    f'<span style="vertical-align: middle;">  </span>'
                    f'<span style="vertical-align: middle; color: #ffffff !important; text-decoration: none !important;">{email_truque}</span>'
                )
            elif email:
                lignes_contact.append(
                    f'<a href="mailto:{email}" style="font-size: 14px; color: #ffffff !important; text-decoration: underline !important; display: inline-block; vertical-align: middle;">✉</a>'
                    f'<span style="vertical-align: middle;">  </span>'
                    f'<span style="vertical-align: middle; color: #ffffff !important; text-decoration: none !important;">{email}</span>'
                )
                
            # 2. ADRESSE
            if adresse:
                adresse_encodee = urllib.parse.quote_plus(adresse)
                lien_maps = f"https://www.google.com/maps/search/?api=1&query={adresse_encodee}"
                lignes_contact.append(
                    f'<a href="{lien_maps}" target="_blank" style="font-size: 18px; font-weight: bold; font-family: Arial, sans-serif; color: #ffffff !important; text-decoration: underline !important; display: inline-block; vertical-align: middle;">⌂</a>'
                    f'<span style="vertical-align: middle;">  </span>'
                    f'<span style="vertical-align: middle; color: #ffffff !important; text-decoration: none !important;">{adresse}</span>'
                )
                
            # 3. TELEPHONE
            if telephone:
                lignes_contact.append(
                    f'<a href="tel:{telephone}" style="font-size: 16px; color: #ffffff !important; text-decoration: underline !important; display: inline-block; vertical-align: middle;">✆</a>'
                    f'<span style="vertical-align: middle;">  </span>'
                    f'<span style="vertical-align: middle; color: #ffffff !important; text-decoration: none !important;">{telephone}</span>'
                )
            
            bloc_contact = "<br>".join(lignes_contact) if lignes_contact else " "
            
            # Attribution sécurisée des couleurs et des logos (triés par taille de texte décroissante)
            couleurs_triees = sorted(COULEURS_ENTREPRISES.items(), key=lambda x: len(x[0]), reverse=True)
            couleur_signature = "#7f8c8d"
            for cle, code_hexa in couleurs_triees:
                if cle in entreprise_brute:
                    couleur_signature = code_hexa
                    break
            
            logos_tries = sorted(LOGOS_ENTREPRISES.items(), key=lambda x: len(x[0]), reverse=True)
            logo_signature = LOGO_PAR_DEFAUT
            for cle, url_logo in logos_tries:
                if cle in entreprise_brute:
                    logo_signature = url_logo
                    break
                    
            html_personnalise = html_template.format(
                logo=logo_signature,
                prenom=prenom_affichage,
                nom=nom_affichage,
                poste=poste,
                bloc_contact=bloc_contact,
                couleur=couleur_signature
            )
            
            # Fichier de sortie
            filename = f"signature_{entreprise_nettoyee}_{nom_clean}_{prenom_clean}.html"
            filepath = os.path.join(output_dir, filename)
            
            with open(filepath, 'w', encoding='utf-8') as out_file:
                out_file.write(html_personnalise)
            
            compteur += 1
            print(f"Signature generee : {prenom_affichage} {nom_affichage} ({entreprise_brute.upper()})")
            
    print(f"\n🎉 Termine ! {compteur} signatures generees à la perfection dans '{output_dir}'.")

except FileNotFoundError:
    print(f"Erreur : Le fichier '{nom_fichier_csv}' est introuvable.")
