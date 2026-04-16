// =============================================================
//  SIGNATURE AUTO ADD-IN - signature.js
//  Récupère le profil M365 via Graph API et injecte la signature
//  dans les nouveaux emails ET les réponses/transferts.
//
//  Prérequis :
//  - App Entra ID enregistrée avec permission déléguée : User.Read
//  - CLIENT_ID à remplacer par l'Application (client) ID de l'app Entra
//  - LOGO_URL à remplacer par l'URL publique du logo hébergé
// =============================================================

Office.initialize = function () {};

// ─── Configuration ────────────────────────────────────────────
const CONFIG = {
  CLIENT_ID: "VOTRE_CLIENT_ID_ENTRA",          // App Registration > Application (client) ID
  LOGO_URL:  "https://VOTRE_DOMAINE.COM/logo.png", // URL publique du logo (remplacer)
  LOGO_HEIGHT: "50",                             // Hauteur du logo en px
  COULEUR_PRINCIPALE: "#003366",                 // Couleur de la barre latérale
};

// ─── Point d'entrée : nouveau message ─────────────────────────
async function insertSignatureOnCompose(event) {
  await insertSignature(event);
}

// ─── Point d'entrée : réponse / transfert ─────────────────────
async function insertSignatureOnReply(event) {
  await insertSignature(event);
}

// ─── Fonction principale ───────────────────────────────────────
async function insertSignature(event) {
  try {
    const token = await getAccessToken();
    const profile = await getUserProfile(token);
    const signatureHTML = buildSignatureHTML(profile);

    Office.context.mailbox.item.body.setSignatureAsync(
      signatureHTML,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Erreur injection signature:", result.error.message);
        }
        event.completed();
      }
    );
  } catch (err) {
    console.error("Erreur add-in signature:", err);
    event.completed();
  }
}

// ─── Obtenir le token SSO via Office.auth ─────────────────────
async function getAccessToken() {
  try {
    // SSO silencieux via Office.auth (aucune popup si l'user est déjà connecté)
    const token = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });
    return token;
  } catch (err) {
    throw new Error("Impossible d'obtenir le token SSO : " + err.message);
  }
}

// ─── Appel Graph API : profil utilisateur ─────────────────────
async function getUserProfile(token) {
  const fields = "displayName,jobTitle,mobilePhone,department,companyName,mail";
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me?$select=${fields}`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
    }
  );

  if (!response.ok) {
    throw new Error(`Graph API erreur ${response.status}: ${response.statusText}`);
  }

  return await response.json();
}

// ─── Construction du HTML de la signature ─────────────────────
function buildSignatureHTML(profile) {
  const nom       = profile.displayName  || "";
  const titre     = profile.jobTitle     || "";
  const telephone = profile.mobilePhone  || "";
  const dept      = profile.department   || "";
  const company   = profile.companyName  || "";
  const email     = profile.mail         || "";

  // Ligne téléphone conditionnelle (masquée si vide dans Entra)
  const lignePhone = telephone
    ? `<tr>
         <td style="padding:1px 0; color:#555555; font-size:12px;">
           📞 <a href="tel:${telephone}" style="color:#555555; text-decoration:none;">${telephone}</a>
         </td>
       </tr>`
    : "";

  // Ligne email conditionnelle
  const ligneEmail = email
    ? `<tr>
         <td style="padding:1px 0; color:#555555; font-size:12px;">
           ✉ <a href="mailto:${email}" style="color:${CONFIG.COULEUR_PRINCIPALE}; text-decoration:none;">${email}</a>
         </td>
       </tr>`
    : "";

  // Ligne département conditionnelle
  const ligneDept = dept
    ? `<tr>
         <td style="padding:1px 0; color:#888888; font-size:11px;">${dept}</td>
       </tr>`
    : "";

  // Bloc logo conditionnel
  const blocLogo = CONFIG.LOGO_URL && !CONFIG.LOGO_URL.includes("VOTRE_DOMAINE")
    ? `<td style="padding-left:16px; vertical-align:middle;">
         <img src="${CONFIG.LOGO_URL}" height="${CONFIG.LOGO_HEIGHT}" alt="${company}" style="display:block;"/>
       </td>`
    : "";

  return `
<table cellpadding="0" cellspacing="0" border="0"
       style="font-family: Calibri, Arial, sans-serif; font-size:13px; color:#222222;
              border-top: 2px solid ${CONFIG.COULEUR_PRINCIPALE}; padding-top:10px; margin-top:10px;">
  <tr>
    <!-- Colonne texte -->
    <td style="vertical-align:top; padding-right:16px;">
      <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td style="font-size:15px; font-weight:bold; color:${CONFIG.COULEUR_PRINCIPALE}; padding-bottom:2px;">
            ${nom}
          </td>
        </tr>
        ${titre ? `<tr><td style="font-size:12px; color:#444444; padding-bottom:4px;">${titre}</td></tr>` : ""}
        ${ligneDept}
        ${lignePhone}
        ${ligneEmail}
        ${company ? `<tr><td style="padding-top:4px; font-size:11px; color:#aaaaaa;">${company}</td></tr>` : ""}
      </table>
    </td>

    <!-- Séparateur vertical -->
    ${blocLogo ? `<td style="border-left:1px solid #dddddd; padding:0;"></td>` : ""}

    <!-- Logo -->
    ${blocLogo}
  </tr>
</table>
`;
}

// ─── Enregistrement des handlers Office ───────────────────────
Office.actions.associate("insertSignatureOnCompose", insertSignatureOnCompose);
Office.actions.associate("insertSignatureOnReply",   insertSignatureOnReply);
