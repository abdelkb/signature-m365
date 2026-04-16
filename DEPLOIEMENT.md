# 📧 Signature Auto Add-in — Guide de déploiement

## Vue d'ensemble

Cet add-in Outlook récupère automatiquement les informations du profil Microsoft 365 (Entra ID) de chaque utilisateur et injecte une signature standardisée dans tous les nouveaux emails, réponses et transferts.

**Données injectées depuis Entra ID :**
- Nom complet (`displayName`)
- Titre / poste (`jobTitle`)
- Téléphone mobile (`mobilePhone`)
- Département (`department`)
- Nom de l'entreprise (`companyName`)
- Adresse courriel (`mail`)
- Logo (URL hébergée publiquement)

---

## Étape 1 — Héberger les fichiers

Les fichiers doivent être accessibles publiquement via HTTPS (Outlook doit pouvoir les charger).

**Options recommandées :**
- Azure Static Web Apps (gratuit)
- GitHub Pages (gratuit)
- SharePoint Online (dossier public)
- Tout hébergeur web avec certificat SSL valide

**Structure à déployer :**
```
/addin/
  ├── commands.html
  ├── signature.js
  └── assets/
        ├── icon-64.png
        └── icon-128.png
```

**Remplacer dans tous les fichiers :** `VOTRE_DOMAINE.COM` → votre domaine réel.

---

## Étape 2 — Enregistrer l'app dans Entra ID

1. Portail Azure → **Entra ID → App registrations → New registration**
2. Nom : `Signature Add-in Outlook`
3. Supported account types : **Accounts in this organizational directory only**
4. Redirect URI : laisser vide pour l'instant → **Register**
5. Dans **Authentication** :
   - Cocher **Access tokens** et **ID tokens**
   - Ajouter platform : **Single-page application** → URI : `https://VOTRE_DOMAINE.COM/addin/commands.html`
6. Dans **API Permissions** :
   - Ajouter permission déléguée : `Microsoft Graph → User.Read`
   - Cliquer **Grant admin consent**
7. Copier l'**Application (client) ID** → le coller dans `signature.js` à la place de `VOTRE_CLIENT_ID_ENTRA`

---

## Étape 3 — Générer un GUID unique pour le manifest

```powershell
[System.Guid]::NewGuid()
```

Remplacer dans `manifest.xml` :
```xml
<Id>NOUVEAU-GUID-ICI</Id>
```

---

## Étape 4 — Ajouter le logo

Héberger le logo de l'entreprise (PNG transparent recommandé, ~200px de large) à une URL publique HTTPS.

Dans `signature.js`, remplacer :
```javascript
LOGO_URL: "https://VOTRE_DOMAINE.COM/logo.png"
```

---

## Étape 5 — Déployer l'add-in via M365 Admin Center

1. **Microsoft 365 Admin Center** → **Settings → Integrated apps**
2. Cliquer **Upload custom apps**
3. Sélectionner **Upload manifest file (.xml)**
4. Uploader `manifest.xml`
5. Choisir les utilisateurs / groupes destinataires (ou **Everyone**)
6. Valider le déploiement

> ⏱️ La propagation peut prendre jusqu'à 24h, mais est généralement immédiate pour Outlook Web.

---

## Étape 6 — S'assurer que les profils Entra sont remplis

L'add-in injecte uniquement les champs renseignés. Les champs vides sont masqués proprement.

Vérifier/remplir en masse via PowerShell :

```powershell
# Connexion
Connect-MgGraph -Scopes "User.ReadWrite.All"

# Voir les users sans titre
Get-MgUser -All -Property DisplayName,JobTitle,MobilePhone |
  Where-Object { -not $_.JobTitle } |
  Select-Object DisplayName, UserPrincipalName

# Mettre à jour un user spécifique
Update-MgUser -UserId "user@domaine.com" `
  -JobTitle "Technicien IT" `
  -MobilePhone "+1 514 555-0000" `
  -Department "Informatique"
```

---

## Personnalisation visuelle

Dans `signature.js`, modifier le bloc `CONFIG` :

```javascript
const CONFIG = {
  CLIENT_ID:          "votre-client-id",
  LOGO_URL:           "https://...",
  LOGO_HEIGHT:        "50",           // px
  COULEUR_PRINCIPALE: "#003366",      // Bleu corporate → adapter à la charte
};
```

---

## Compatibilité

| Client          | Support           |
|-----------------|-------------------|
| Outlook Web     | ✅ Complet         |
| Outlook Desktop | ✅ Windows & Mac   |
| Outlook Mobile  | ⚠️ Partiel (iOS/Android — setSignatureAsync limité) |
| Nouveau Outlook | ✅ Complet         |

---

## Dépannage

| Problème | Solution |
|----------|----------|
| Signature n'apparaît pas | Vérifier la propagation (attendre 24h) / vérifier la console F12 |
| Erreur token SSO | Vérifier les Redirect URIs dans l'app Entra + admin consent |
| Champs vides dans la signature | Remplir les attributs dans Entra ID / Admin Center M365 |
| Logo ne s'affiche pas | Vérifier que l'URL est publique et en HTTPS |
| Erreur `setSignatureAsync` | Vérifier que le manifest déclare `Mailbox 1.3` minimum |
