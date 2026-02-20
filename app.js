// Configuration for MSAL (Microsoft Authentication Library)
const msalConfig = {
  auth: {
    clientId: "58d4f2d3-5598-401e-a2ff-a01806d304e7",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "http://localhost:8080"
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

// Permissions we need
const loginRequest = {
  scopes: ["User.Read", "Mail.Read", "Calendars.Read", "Files.Read.All"]
};

// ---- Environment detection ----
const IS_GITHUB_PAGES = location.hostname.includes("github.io");

// ---- i18n (ES / EN / DE) ----
const LANG_DEFAULT = "es"; // ES, EN, DE (in this order)

const i18n = {
  es: {
    title: "🚀 My Microsoft 365 Dashboard",
    subtitle: "Conecta con tu cuenta de Microsoft y explora tus datos",
    login: "🔐 Iniciar Sesión con Microsoft",
    logout: "Cerrar Sesión",
    welcome: "¡Bienvenido!",
    profileTitle: "👤 Mi Perfil",
    emailsTitle: "📧 Últimos Emails",
    calendarTitle: "📅 Próximos Eventos",
    filesTitle: "📁 Archivos Recientes",
    loadingProfile: "Cargando perfil...",
    loadingEmails: "Cargando emails...",
    loadingCalendar: "Cargando calendario...",
    loadingFiles: "Cargando archivos...",
    footer1: "💡 Proyecto de demostración - Microsoft Graph API",
    footer2: "Usando MSAL.js 2.0 para autenticación",

    name: "Nombre",
    email: "Email",
    jobTitle: "Puesto",
    location: "Ubicación",
    from: "De",
    modified: "Modificado",
    unknown: "Desconocido",
    na: "N/A",

    noEmails: "No hay emails recientes",
    noFiles: "No hay archivos recientes",
    noEvents7: "No hay eventos próximos en los próximos 7 días",
    noAuthUser: "No hay usuario autenticado",

    errLogin: "Error al iniciar sesión",
    errLoad: "Error al cargar los datos",
    errProfile: "❌ Error al cargar el perfil",
    errEmails: "❌ Error al cargar emails (puede que no tengas buzón configurado)",
    errCalendar: "❌ Error al cargar calendario",
    errFiles: "❌ Error al cargar archivos (puede que no tengas OneDrive configurado)",

    noSubject: "(Sin asunto)",
    noTitle: "(Sin título)",
    userFallback: "Usuario"
  
,
    demoButton: "👀 Ver demo (sin login)",
    demoBanner: "Estás viendo el modo demo. Inicia sesión para ver tus datos reales.",
    demoUserName: "Usuario Demo",
    demoJobTitle: "Cloud / IT Support (Demo)",
    demoLocation: "Basel (Demo)",
    demoMailSubject1: "Documentación del laboratorio (GitHub)",
    demoMailSubject2: "Tu suscripción Microsoft 365 (Demo)",
    demoMailSubject3: "Checklist para entrevista técnica",
    demoEvent1: "Revisión de redes (Demo)",
    demoEvent2: "Prep entrevista: Azure / M365 (Demo)",
    demoEventLoc1: "Teams (Demo)",
    demoEventLoc2: "Oficina (Demo)"
  },
  en: {
    title: "🚀 My Microsoft 365 Dashboard",
    subtitle: "Connect with your Microsoft account and explore your data",
    login: "🔐 Sign in with Microsoft",
    logout: "Sign out",
    welcome: "Welcome!",
    profileTitle: "👤 My Profile",
    emailsTitle: "📧 Latest Emails",
    calendarTitle: "📅 Upcoming Events",
    filesTitle: "📁 Recent Files",
    loadingProfile: "Loading profile...",
    loadingEmails: "Loading emails...",
    loadingCalendar: "Loading calendar...",
    loadingFiles: "Loading files...",
    footer1: "💡 Demo project - Microsoft Graph API",
    footer2: "Using MSAL.js 2.0 for authentication",

    name: "Name",
    email: "Email",
    jobTitle: "Job Title",
    location: "Location",
    from: "From",
    modified: "Modified",
    unknown: "Unknown",
    na: "N/A",

    noEmails: "No recent emails",
    noFiles: "No recent files",
    noEvents7: "No upcoming events in the next 7 days",
    noAuthUser: "No authenticated user",

    errLogin: "Login error",
    errLoad: "Error loading data",
    errProfile: "❌ Error loading profile",
    errEmails: "❌ Error loading emails (you may not have a mailbox configured)",
    errCalendar: "❌ Error loading calendar",
    errFiles: "❌ Error loading files (you may not have OneDrive configured)",

    noSubject: "(No subject)",
    noTitle: "(No title)",
    userFallback: "User"
  
,
    demoButton: "👀 View demo (no login)",
    demoBanner: "You are viewing demo mode. Sign in to see your real data.",
    demoUserName: "Demo User",
    demoJobTitle: "Cloud / IT Support (Demo)",
    demoLocation: "Basel (Demo)",
    demoMailSubject1: "Lab documentation (GitHub)",
    demoMailSubject2: "Your Microsoft 365 subscription (Demo)",
    demoMailSubject3: "Technical interview checklist",
    demoEvent1: "Network review (Demo)",
    demoEvent2: "Interview prep: Azure / M365 (Demo)",
    demoEventLoc1: "Teams (Demo)",
    demoEventLoc2: "Office (Demo)"
  },
  de: {
    title: "🚀 My Microsoft 365 Dashboard",
    subtitle: "Melde dich mit deinem Microsoft-Konto an und erkunde deine Daten",
    login: "🔐 Mit Microsoft anmelden",
    logout: "Abmelden",
    welcome: "Willkommen!",
    profileTitle: "👤 Mein Profil",
    emailsTitle: "📧 Letzte E-Mails",
    calendarTitle: "📅 Nächste Termine",
    filesTitle: "📁 Letzte Dateien",
    loadingProfile: "Profil wird geladen...",
    loadingEmails: "E-Mails werden geladen...",
    loadingCalendar: "Kalender wird geladen...",
    loadingFiles: "Dateien werden geladen...",
    footer1: "💡 Demo-Projekt - Microsoft Graph API",
    footer2: "Mit MSAL.js 2.0 zur Authentifizierung",

    name: "Name",
    email: "E-Mail",
    jobTitle: "Position",
    location: "Standort",
    from: "Von",
    modified: "Geändert",
    unknown: "Unbekannt",
    na: "k. A.",

    noEmails: "Keine aktuellen E-Mails",
    noFiles: "Keine aktuellen Dateien",
    noEvents7: "Keine Termine in den nächsten 7 Tagen",
    noAuthUser: "Kein authentifizierter Benutzer",

    errLogin: "Anmeldefehler",
    errLoad: "Fehler beim Laden der Daten",
    errProfile: "❌ Fehler beim Laden des Profils",
    errEmails: "❌ Fehler beim Laden der E-Mails (evtl. ist kein Postfach eingerichtet)",
    errCalendar: "❌ Fehler beim Laden des Kalenders",
    errFiles: "❌ Fehler beim Laden der Dateien (evtl. ist OneDrive nicht eingerichtet)",

    noSubject: "(Kein Betreff)",
    noTitle: "(Kein Titel)",
    userFallback: "Benutzer"
  
,
    demoButton: "👀 Demo ansehen (ohne Login)",
    demoBanner: "Du siehst den Demo-Modus. Melde dich an, um deine echten Daten zu sehen.",
    demoUserName: "Demo-Benutzer",
    demoJobTitle: "Cloud / IT-Support (Demo)",
    demoLocation: "Basel (Demo)",
    demoMailSubject1: "Lab-Dokumentation (GitHub)",
    demoMailSubject2: "Dein Microsoft 365 Abonnement (Demo)",
    demoMailSubject3: "Checkliste fürs Tech-Interview",
    demoEvent1: "Netzwerk-Review (Demo)",
    demoEvent2: "Interview-Vorbereitung: Azure / M365 (Demo)",
    demoEventLoc1: "Teams (Demo)",
    demoEventLoc2: "Büro (Demo)"
  }
};

function getLang() {
  return localStorage.getItem("lang") || LANG_DEFAULT;
}
function setLang(lang) {
  localStorage.setItem("lang", lang);
  document.documentElement.lang = getLocale(); // es-ES/en-US/de-DE
}
function t(key) {
  const lang = getLang();
  return (i18n[lang] && i18n[lang][key]) ? i18n[lang][key] : (i18n[LANG_DEFAULT][key] || key);
}
function getLocale() {
  const lang = getLang();
  if (lang === "de") return "de-DE";
  if (lang === "en") return "en-US";
  return "es-ES";
}

// ---- Initialize MSAL ----
const msalInstance = new msal.PublicClientApplication(msalConfig);

// ---- App state (for re-render without refetch) ----
const state = { profile: null, emails: null, events: null, files: null };
let demoMode = false;

// ---- DOM refs (initialized on DOMContentLoaded) ----
let loginButton, logoutButton, demoButton, loginSection, content, userName;
let demoBanner;
let profileSection, emailSection, calendarSection, filesSection;
let appTitleEl, appSubtitleEl, welcomeTextEl, profileTitleEl, emailTitleEl, calendarTitleEl, filesTitleEl, footer1El, footer2El, langSelect;

// ---- DOM Ready ----
document.addEventListener("DOMContentLoaded", () => {
  // Bind elements safely
  loginButton = document.getElementById("login-button");
  logoutButton = document.getElementById("logout-button");
  demoButton = document.getElementById("demo-button");
  demoBanner = document.getElementById("demo-banner");
  loginSection = document.getElementById("login-section");
  content = document.getElementById("content");
  userName = document.getElementById("user-name");

  profileSection = document.getElementById("profile-section");
  emailSection = document.getElementById("email-section");
  calendarSection = document.getElementById("calendar-section");
  filesSection = document.getElementById("files-section");

  appTitleEl = document.getElementById("app-title");
  appSubtitleEl = document.getElementById("app-subtitle");
  welcomeTextEl = document.getElementById("welcome-text");
  profileTitleEl = document.getElementById("profile-title");
  emailTitleEl = document.getElementById("email-title");
  calendarTitleEl = document.getElementById("calendar-title");
  filesTitleEl = document.getElementById("files-title");
  footer1El = document.getElementById("footer-line1");
  footer2El = document.getElementById("footer-line2");
  langSelect = document.getElementById("lang-select");

  // Language init
  const currentLang = getLang();
  if (langSelect) {
    langSelect.value = currentLang;
    langSelect.addEventListener("change", () => {
      setLang(langSelect.value);
      applyTranslations();
      rerenderAll();
    });
  }
  setLang(currentLang);
  applyTranslations();

  // Events (only when DOM exists!)
  loginButton?.addEventListener("click", login);
  logoutButton?.addEventListener("click", logout);
  demoButton?.addEventListener("click", startDemo);

  // Check session AFTER everything is ready
  if (IS_GITHUB_PAGES) {
    // Auto-demo on GitHub Pages (no login required)
    startDemo();
  } else {
    checkAccount();
  }
});
// ---- UI helpers ----
function showContent() {
  loginSection.style.display = "none";
  content.style.display = "block";
}
function hideContent() {
  loginSection.style.display = "block";
  content.style.display = "none";
}

function applyTranslations() {
  document.title = "My Microsoft 365 Dashboard";

  appTitleEl && (appTitleEl.textContent = t("title"));
  appSubtitleEl && (appSubtitleEl.textContent = t("subtitle"));
  loginButton && (loginButton.textContent = t("login"));
  logoutButton && (logoutButton.textContent = t("logout"));
  demoButton && (demoButton.textContent = t("demoButton"));
  if (demoBanner && demoBanner.style.display !== "none") demoBanner.textContent = t("demoBanner");
  welcomeTextEl && (welcomeTextEl.textContent = t("welcome"));

  profileTitleEl && (profileTitleEl.textContent = t("profileTitle"));
  emailTitleEl && (emailTitleEl.textContent = t("emailsTitle"));
  calendarTitleEl && (calendarTitleEl.textContent = t("calendarTitle"));
  filesTitleEl && (filesTitleEl.textContent = t("filesTitle"));

  footer1El && (footer1El.textContent = t("footer1"));
  footer2El && (footer2El.textContent = t("footer2"));

  if (!state.profile && profileSection) profileSection.textContent = t("loadingProfile");
  if (!state.emails && emailSection) emailSection.textContent = t("loadingEmails");
  if (!state.events && calendarSection) calendarSection.textContent = t("loadingCalendar");
  if (!state.files && filesSection) filesSection.textContent = t("loadingFiles");
}

function rerenderAll() {
  if (state.profile) renderProfile(state.profile);
  if (state.emails) renderEmails(state.emails);
  if (state.events) renderCalendar(state.events);
  if (state.files) renderFiles(state.files);
}

// ---- Auth ----
async function checkAccount() {
  const currentAccounts = msalInstance.getAllAccounts();
  if (currentAccounts.length > 0) {
    // Ensure texts (logout button etc.) are applied even on auto-login
    demoMode = false;
    if (demoBanner) demoBanner.style.display = "none";
    applyTranslations();
    showContent();
    loadUserData();
  }
}


function startDemo() {
  demoMode = true;

  // Show dashboard without authentication
  showContent();

  // Demo banner
  if (demoBanner) {
    demoBanner.style.display = "block";
    demoBanner.textContent = t("demoBanner");
  }

  // Build demo data
  const now = new Date();
  const plusHours = (h) => new Date(now.getTime() + h * 60 * 60 * 1000).toISOString();

  const demoProfile = {
    displayName: t("demoUserName"),
    mail: "demo.user@example.com",
    userPrincipalName: "demo.user@example.com",
    jobTitle: t("demoJobTitle"),
    officeLocation: t("demoLocation")
  };

  const demoEmails = [
    {
      subject: t("demoMailSubject1"),
      from: { emailAddress: { name: "Contoso HR" } },
      receivedDateTime: plusHours(-2),
      isRead: false
    },
    {
      subject: t("demoMailSubject2"),
      from: { emailAddress: { name: "Microsoft 365" } },
      receivedDateTime: plusHours(-6),
      isRead: true
    },
    {
      subject: t("demoMailSubject3"),
      from: { emailAddress: { name: "Team Lead" } },
      receivedDateTime: plusHours(-20),
      isRead: true
    }
  ];

  const demoEvents = [
    {
      subject: t("demoEvent1"),
      start: { dateTime: plusHours(6) },
      end: { dateTime: plusHours(7) },
      location: { displayName: t("demoEventLoc1") }
    },
    {
      subject: t("demoEvent2"),
      start: { dateTime: plusHours(30) },
      end: { dateTime: plusHours(31) },
      location: { displayName: t("demoEventLoc2") }
    }
  ];

  const demoFiles = [
    { name: "CV_Vidal_Renao.pdf", lastModifiedDateTime: plusHours(-12), size: 352000 },
    { name: "Azure-Arc-Lab-Notes.docx", lastModifiedDateTime: plusHours(-28), size: 118000 },
    { name: "Network-Diagram.png", lastModifiedDateTime: plusHours(-40), size: 890000 }
  ];

  state.profile = demoProfile;
  state.emails = demoEmails;
  state.events = demoEvents;
  state.files = demoFiles;

  applyTranslations();
  rerenderAll();
}

async function login() {
  if (IS_GITHUB_PAGES) {
    // On GitHub Pages we run in demo mode to avoid redirect URI issues.
    startDemo();
    return;
  }
  try {
    const loginResponse = await msalInstance.loginPopup(loginRequest);
    console.log("Login successful:", loginResponse);
    showContent();
    loadUserData();
  } catch (error) {
    console.error("Login error:", error);
    alert(`${t("errLogin")}: ${error.message}`);
  }
}

function logout() {
  demoMode = false;
  if (demoBanner) demoBanner.style.display = "none";
  const currentAccounts = msalInstance.getAllAccounts();
  if (currentAccounts.length > 0) {
    msalInstance.logoutPopup({ account: currentAccounts[0] });
  }

  // reset state
  state.profile = null;
  state.emails = null;
  state.events = null;
  state.files = null;

  if (userName) userName.textContent = "";
  applyTranslations();
  hideContent();
}

async function getAccessToken() {
  const currentAccounts = msalInstance.getAllAccounts();
  if (currentAccounts.length === 0) {
    throw new Error(t("noAuthUser"));
  }

  const request = { scopes: loginRequest.scopes, account: currentAccounts[0] };

  try {
    const response = await msalInstance.acquireTokenSilent(request);
    return response.accessToken;
  } catch (error) {
    console.log("Silent token acquisition failed, acquiring token using popup");
    const response = await msalInstance.acquireTokenPopup(request);
    return response.accessToken;
  }
}

// ---- Data loading ----
async function loadUserData() {
  if (demoMode) return;
  try {
    const accessToken = await getAccessToken();

    await Promise.all([
      loadProfile(accessToken),
      loadEmails(accessToken),
      loadCalendar(accessToken),
      loadFiles(accessToken)
    ]);
  } catch (error) {
    console.error("Error loading user data:", error);
    alert(`${t("errLoad")}: ${error.message}`);
  }
}

async function loadProfile(accessToken) {
  try {
    const response = await fetch("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

    const profile = await response.json();
    state.profile = profile;
    renderProfile(profile);
  } catch (error) {
    console.error("Error loading profile:", error);
    profileSection.innerHTML = `<p class="error">${t("errProfile")}</p>`;
  }
}

function renderProfile(profile) {
  const displayName = profile.displayName || t("userFallback");
  if (userName) userName.textContent = displayName;

  profileSection.innerHTML = `
    <div class="profile-info">
      <p><strong>${t("name")}:</strong> ${profile.displayName || t("na")}</p>
      <p><strong>${t("email")}:</strong> ${profile.mail || profile.userPrincipalName || t("na")}</p>
      <p><strong>${t("jobTitle")}:</strong> ${profile.jobTitle || t("na")}</p>
      <p><strong>${t("location")}:</strong> ${profile.officeLocation || t("na")}</p>
    </div>
  `;
}

async function loadEmails(accessToken) {
  try {
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/messages?$top=5&$select=subject,from,receivedDateTime,isRead",
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

    const data = await response.json();
    const emails = data.value || [];
    state.emails = emails;
    renderEmails(emails);
  } catch (error) {
    console.error("Error loading emails:", error);
    emailSection.innerHTML = `<p class="error">${t("errEmails")}</p>`;
  }
}

function renderEmails(emails) {
  if (!emails || emails.length === 0) {
    emailSection.innerHTML = `<p>${t("noEmails")}</p>`;
    return;
  }

  const locale = getLocale();
  let emailHTML = '<div class="email-list">';
  emails.forEach((email) => {
    const date = new Date(email.receivedDateTime).toLocaleString(locale);
    const readClass = email.isRead ? "read" : "unread";
    emailHTML += `
      <div class="email-item ${readClass}">
        <div class="email-subject">${email.subject || t("noSubject")}</div>
        <div class="email-from">${t("from")}: ${email.from?.emailAddress?.name || t("unknown")}</div>
        <div class="email-date">${date}</div>
      </div>
    `;
  });
  emailHTML += "</div>";
  emailSection.innerHTML = emailHTML;
}

async function loadCalendar(accessToken) {
  try {
    const now = new Date().toISOString();
    const end = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${now}&endDateTime=${end}&$top=5&$select=subject,start,end,location`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

    const data = await response.json();
    const events = data.value || [];
    state.events = events;
    renderCalendar(events);
  } catch (error) {
    console.error("Error loading calendar:", error);
    calendarSection.innerHTML = `<p class="error">${t("errCalendar")}</p>`;
  }
}

function renderCalendar(events) {
  if (!events || events.length === 0) {
    calendarSection.innerHTML = `<p>${t("noEvents7")}</p>`;
    return;
  }

  const locale = getLocale();
  let calendarHTML = '<div class="calendar-list">';
  events.forEach((event) => {
    const startDate = new Date(event.start.dateTime).toLocaleString(locale);
    calendarHTML += `
      <div class="calendar-item">
        <div class="event-subject">${event.subject || t("noTitle")}</div>
        <div class="event-time">📅 ${startDate}</div>
        ${
          event.location?.displayName
            ? `<div class="event-location">📍 ${event.location.displayName}</div>`
            : ""
        }
      </div>
    `;
  });
  calendarHTML += "</div>";
  calendarSection.innerHTML = calendarHTML;
}

async function loadFiles(accessToken) {
  try {
    const response = await fetch("https://graph.microsoft.com/v1.0/me/drive/recent?$top=5", {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

    const data = await response.json();
    const files = data.value || [];
    state.files = files;
    renderFiles(files);
  } catch (error) {
    console.error("Error loading files:", error);
    filesSection.innerHTML = `<p class="error">${t("errFiles")}</p>`;
  }
}

function renderFiles(files) {
  if (!files || files.length === 0) {
    filesSection.innerHTML = `<p>${t("noFiles")}</p>`;
    return;
  }

  const locale = getLocale();
  let filesHTML = '<div class="files-list">';
  files.forEach((file) => {
    const modifiedDate = new Date(file.lastModifiedDateTime).toLocaleString(locale);
    const size = formatFileSize(file.size);
    const icon = getFileIcon(file.name);

    filesHTML += `
      <div class="file-item">
        <div class="file-icon">${icon}</div>
        <div class="file-info">
          <div class="file-name">${file.name}</div>
          <div class="file-details">${size} • ${t("modified")}: ${modifiedDate}</div>
        </div>
      </div>
    `;
  });
  filesHTML += "</div>";
  filesSection.innerHTML = filesHTML;
}

// ---- Utilities ----
function formatFileSize(bytes) {
  if (!bytes || bytes === 0) return "0 Bytes";
  const k = 1024;
  const sizes = ["Bytes", "KB", "MB", "GB", "TB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + " " + sizes[i];
}

function getFileIcon(filename) {
  const ext = filename.split(".").pop().toLowerCase();
  const icons = {
    pdf: "📄",
    doc: "📝",
    docx: "📝",
    xls: "📊",
    xlsx: "📊",
    ppt: "📊",
    pptx: "📊",
    txt: "📃",
    jpg: "🖼️",
    jpeg: "🖼️",
    png: "🖼️",
    gif: "🖼️",
    zip: "📦",
    rar: "📦"
  };
  return icons[ext] || "📄";
}
