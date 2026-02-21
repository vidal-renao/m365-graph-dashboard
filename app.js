// Microsoft 365 Dashboard (MSAL + Microsoft Graph) - ES/EN/DE + Quick Stats + GitHub Pages demo mode

// -------------------- MSAL configuration --------------------
const msalConfig = {
  auth: {
    clientId: "58d4f2d3-5598-401e-a2ff-a01806d304e7",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "http://localhost:8080" // local dev (GitHub Pages runs demo mode)
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

// Environment detection
const IS_GITHUB_PAGES = location.hostname.includes("github.io");

// Initialize MSAL
const msalInstance = new msal.PublicClientApplication(msalConfig);

// -------------------- Guard: avoid running full app inside MSAL popup --------------------
// With loginPopup/acquireTokenPopup, MSAL opens a popup and navigates it.
// We must not run our whole dashboard inside that popup (MSAL blocks nested popups).
const IS_MSAL_POPUP =
  !!window.opener &&
  (String(window.name || "").toLowerCase().includes("msal") || location.hash.includes("code="));

if (IS_MSAL_POPUP) {
  // Let MSAL finish processing then close the popup.
  msalInstance
    .handleRedirectPromise()
    .finally(() => window.close());
  // Stop here.
} else {

// -------------------- i18n --------------------
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

    // labels
    name: "Nombre",
    email: "Email",
    jobTitle: "Puesto",
    location: "Ubicación",
    from: "De",
    modified: "Modificado",
    unknown: "Desconocido",
    na: "N/A",

    // empty states
    noEmails: "No hay emails recientes",
    noFiles: "No hay archivos recientes",
    noEvents7: "No hay eventos próximos en los próximos 7 días",
    noAuthUser: "No hay usuario autenticado",
    none: "Ninguno",

    // errors
    errLogin: "Error al iniciar sesión",
    errLoad: "Error al cargar los datos",
    errProfile: "❌ Error al cargar el perfil",
    errEmails: "❌ Error al cargar emails (puede que no tengas buzón configurado)",
    errCalendar: "❌ Error al cargar calendario",
    errFiles: "❌ Error al cargar archivos (puede que no tengas OneDrive configurado)",

    // fallbacks
    noSubject: "(Sin asunto)",
    noTitle: "(Sin título)",
    userFallback: "Usuario",

    // demo
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
    demoEventLoc2: "Oficina (Demo)",

    // quick stats
    statsTitle: "📊 Estadísticas rápidas",
    statUnread: "Correos sin leer",
    statNextEvent: "Próximo evento",
    statRecentFiles: "Archivos recientes (48h)",
    statUpdated: "Última actualización"
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
    none: "None",

    errLogin: "Login error",
    errLoad: "Error loading data",
    errProfile: "❌ Error loading profile",
    errEmails: "❌ Error loading emails (you may not have a mailbox configured)",
    errCalendar: "❌ Error loading calendar",
    errFiles: "❌ Error loading files (you may not have OneDrive configured)",

    noSubject: "(No subject)",
    noTitle: "(No title)",
    userFallback: "User",

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
    demoEventLoc2: "Office (Demo)",

    statsTitle: "📊 Quick Stats",
    statUnread: "Unread emails",
    statNextEvent: "Next event",
    statRecentFiles: "Recent files (48h)",
    statUpdated: "Last updated"
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
    none: "Keine",

    errLogin: "Anmeldefehler",
    errLoad: "Fehler beim Laden der Daten",
    errProfile: "❌ Fehler beim Laden des Profils",
    errEmails: "❌ Fehler beim Laden der E-Mails (evtl. ist kein Postfach eingerichtet)",
    errCalendar: "❌ Fehler beim Laden des Kalenders",
    errFiles: "❌ Fehler beim Laden der Dateien (evtl. ist OneDrive nicht eingerichtet)",

    noSubject: "(Kein Betreff)",
    noTitle: "(Kein Titel)",
    userFallback: "Benutzer",

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
    demoEventLoc2: "Büro (Demo)",

    statsTitle: "📊 Schnellübersicht",
    statUnread: "Ungelesene E-Mails",
    statNextEvent: "Nächster Termin",
    statRecentFiles: "Letzte Dateien (48 Std.)",
    statUpdated: "Zuletzt aktualisiert"
  }
};

function getLang() {
  return localStorage.getItem("lang") || LANG_DEFAULT;
}

function setLang(lang) {
  localStorage.setItem("lang", lang);
  document.documentElement.lang = lang;
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

// -------------------- State --------------------
const state = {
  profile: null,
  emails: null,
  events: null,
  files: null,
  _lastUpdated: null
};

let demoMode = false;

// -------------------- DOM refs --------------------
let loginButton, logoutButton, demoButton, loginSection, content, userName;
let demoBanner;
let profileSection, emailSection, calendarSection, filesSection;

let appTitleEl, appSubtitleEl, welcomeTextEl, profileTitleEl, emailTitleEl, calendarTitleEl, filesTitleEl, footer1El, footer2El, langSelect;

// Quick stats DOM refs
let statsTitleEl, statUnreadLabelEl, statNextEventLabelEl, statRecentFilesLabelEl, statUpdatedLabelEl;
let statUnreadValueEl, statNextEventValueEl, statRecentFilesValueEl, statUpdatedValueEl;

// -------------------- DOM Ready --------------------
document.addEventListener("DOMContentLoaded", () => {
  // Bind elements
  loginButton = document.getElementById("login-button");
  logoutButton = document.getElementById("logout-button");
  demoButton = document.getElementById("demo-button");

  loginSection = document.getElementById("login-section");
  content = document.getElementById("content");
  userName = document.getElementById("user-name");

  demoBanner = document.getElementById("demo-banner");

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

  statsTitleEl = document.getElementById("stats-title");
  statUnreadLabelEl = document.getElementById("stat-unread-label");
  statNextEventLabelEl = document.getElementById("stat-next-event-label");
  statRecentFilesLabelEl = document.getElementById("stat-recent-files-label");
  statUpdatedLabelEl = document.getElementById("stat-updated-label");

  statUnreadValueEl = document.getElementById("stat-unread-value");
  statNextEventValueEl = document.getElementById("stat-next-event-value");
  statRecentFilesValueEl = document.getElementById("stat-recent-files-value");
  statUpdatedValueEl = document.getElementById("stat-updated-value");

  // Language selector init
  const currentLang = getLang();
  if (langSelect) {
    langSelect.value = currentLang;
    langSelect.addEventListener("change", () => {
      setLang(langSelect.value);
      applyTranslations();
      rerenderAll();
      renderStats();
    });
  }
  setLang(currentLang);

  // Events
  loginButton?.addEventListener("click", login);
  logoutButton?.addEventListener("click", logout);
  demoButton?.addEventListener("click", startDemo);

  // Initial render
  applyTranslations();
  hideContent();
  renderStats();

  // Start
  if (IS_GITHUB_PAGES) {
    // Auto-demo on GitHub Pages to avoid redirect URI config
    startDemo();
  } else {
    checkAccount();
  }
});

// -------------------- UI helpers --------------------
function showContent() {
  if (loginSection) loginSection.style.display = "none";
  if (content) content.style.display = "block";
}

function hideContent() {
  if (loginSection) loginSection.style.display = "block";
  if (content) content.style.display = "none";
}

function applyTranslations() {
  document.title = "My Microsoft 365 Dashboard";

  if (appTitleEl) appTitleEl.textContent = t("title");
  if (appSubtitleEl) appSubtitleEl.textContent = t("subtitle");

  if (loginButton) loginButton.textContent = t("login");
  if (logoutButton) logoutButton.textContent = t("logout");
  if (demoButton) demoButton.textContent = t("demoButton");

  if (welcomeTextEl) welcomeTextEl.textContent = t("welcome");

  if (statsTitleEl) statsTitleEl.textContent = t("statsTitle");
  if (statUnreadLabelEl) statUnreadLabelEl.textContent = t("statUnread");
  if (statNextEventLabelEl) statNextEventLabelEl.textContent = t("statNextEvent");
  if (statRecentFilesLabelEl) statRecentFilesLabelEl.textContent = t("statRecentFiles");
  if (statUpdatedLabelEl) statUpdatedLabelEl.textContent = t("statUpdated");

  if (demoBanner && demoBanner.style.display !== "none") {
    demoBanner.textContent = t("demoBanner");
  }

  if (profileTitleEl) profileTitleEl.textContent = t("profileTitle");
  if (emailTitleEl) emailTitleEl.textContent = t("emailsTitle");
  if (calendarTitleEl) calendarTitleEl.textContent = t("calendarTitle");
  if (filesTitleEl) filesTitleEl.textContent = t("filesTitle");

  if (footer1El) footer1El.textContent = t("footer1");
  if (footer2El) footer2El.textContent = t("footer2");

  // Placeholders only if not loaded yet
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

function renderStats() {
  const locale = getLocale();

  // 1) Unread emails
  const emails = state.emails || [];
  const unreadCount = emails.filter(e => e && e.isRead === false).length;
  if (statUnreadValueEl) statUnreadValueEl.textContent = state.emails ? String(unreadCount) : "—";

  // 2) Next event
  const events = (state.events || []).slice().sort((a, b) => {
    const da = new Date(a?.start?.dateTime || 0).getTime();
    const db = new Date(b?.start?.dateTime || 0).getTime();
    return da - db;
  });
  const next = events[0];
  const nextText = next?.start?.dateTime
    ? new Date(next.start.dateTime).toLocaleString(locale)
    : (state.events ? t("none") : "—");
  if (statNextEventValueEl) statNextEventValueEl.textContent = nextText;

  // 3) Recent files (last 48h)
  const files = state.files || [];
  const now = Date.now();
  const recent48h = files.filter(f => {
    const dt = f?.lastModifiedDateTime ? new Date(f.lastModifiedDateTime).getTime() : 0;
    return dt && (now - dt) <= (48 * 60 * 60 * 1000);
  }).length;
  if (statRecentFilesValueEl) statRecentFilesValueEl.textContent = state.files ? String(recent48h) : "—";

  // 4) Last updated
  if (statUpdatedValueEl) {
    statUpdatedValueEl.textContent = state._lastUpdated
      ? new Date(state._lastUpdated).toLocaleString(locale)
      : "—";
  }
}

// -------------------- Auth --------------------
async function checkAccount() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    demoMode = false;
    if (demoBanner) demoBanner.style.display = "none";
    applyTranslations();
    showContent();
    await loadUserData();
  }
}

async function login() {
  if (IS_GITHUB_PAGES) {
    // Pages: demo mode by default.
    startDemo();
    return;
  }

  try {
    await msalInstance.loginPopup(loginRequest);
    demoMode = false;
    if (demoBanner) demoBanner.style.display = "none";
    applyTranslations();
    showContent();
    await loadUserData();
  } catch (error) {
    console.error("Login error:", error);
    alert(`${t("errLogin")}: ${error.message}`);
  }
}

function logout() {
  // Demo mode: just reset UI
  if (demoMode) {
    demoMode = false;
    if (demoBanner) demoBanner.style.display = "none";
    Object.assign(state, { profile: null, emails: null, events: null, files: null, _lastUpdated: null });
    if (userName) userName.textContent = "";
    applyTranslations();
    hideContent();
    renderStats();
    return;
  }

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    msalInstance.logoutPopup({ account: accounts[0] });
  }

  Object.assign(state, { profile: null, emails: null, events: null, files: null, _lastUpdated: null });
  if (userName) userName.textContent = "";
  applyTranslations();
  hideContent();
  renderStats();
}

async function getAccessToken() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) throw new Error(t("noAuthUser"));

  const request = { scopes: loginRequest.scopes, account: accounts[0] };

  try {
    const response = await msalInstance.acquireTokenSilent(request);
    return response.accessToken;
  } catch (error) {
    const response = await msalInstance.acquireTokenPopup(request);
    return response.accessToken;
  }
}

// -------------------- Data loading --------------------
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

    state._lastUpdated = new Date().toISOString();
    renderStats();
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

    state.profile = await response.json();
    renderProfile(state.profile);
  } catch (error) {
    console.error("Error loading profile:", error);
    if (profileSection) profileSection.innerHTML = `<p class="error">${t("errProfile")}</p>`;
  }
}

function renderProfile(profile) {
  const displayName = profile.displayName || t("userFallback");
  if (userName) userName.textContent = displayName;

  if (!profileSection) return;
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
    state.emails = data.value || [];
    renderEmails(state.emails);
  } catch (error) {
    console.error("Error loading emails:", error);
    if (emailSection) emailSection.innerHTML = `<p class="error">${t("errEmails")}</p>`;
  }
}

function renderEmails(emails) {
  if (!emailSection) return;
  if (!emails || emails.length === 0) {
    emailSection.innerHTML = `<p>${t("noEmails")}</p>`;
    return;
  }

  const locale = getLocale();
  let html = '<div class="email-list">';
  emails.forEach(email => {
    const date = new Date(email.receivedDateTime).toLocaleString(locale);
    const readClass = email.isRead ? "read" : "unread";
    html += `
      <div class="email-item ${readClass}">
        <div class="email-subject">${email.subject || t("noSubject")}</div>
        <div class="email-from">${t("from")}: ${email.from?.emailAddress?.name || t("unknown")}</div>
        <div class="email-date">${date}</div>
      </div>
    `;
  });
  html += "</div>";
  emailSection.innerHTML = html;

  renderStats(); // update unread count
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
    state.events = data.value || [];
    renderCalendar(state.events);
  } catch (error) {
    console.error("Error loading calendar:", error);
    if (calendarSection) calendarSection.innerHTML = `<p class="error">${t("errCalendar")}</p>`;
  }
}

function renderCalendar(events) {
  if (!calendarSection) return;
  if (!events || events.length === 0) {
    calendarSection.innerHTML = `<p>${t("noEvents7")}</p>`;
    renderStats();
    return;
  }

  const locale = getLocale();
  let html = '<div class="calendar-list">';
  events.forEach(event => {
    const startDate = new Date(event.start.dateTime).toLocaleString(locale);
    html += `
      <div class="calendar-item">
        <div class="event-subject">${event.subject || t("noTitle")}</div>
        <div class="event-time">📅 ${startDate}</div>
        ${event.location?.displayName ? `<div class="event-location">📍 ${event.location.displayName}</div>` : ""}
      </div>
    `;
  });
  html += "</div>";
  calendarSection.innerHTML = html;

  renderStats();
}

async function loadFiles(accessToken) {
  try {
    const response = await fetch("https://graph.microsoft.com/v1.0/me/drive/recent?$top=5", {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

    const data = await response.json();
    state.files = data.value || [];
    renderFiles(state.files);
  } catch (error) {
    console.error("Error loading files:", error);
    if (filesSection) filesSection.innerHTML = `<p class="error">${t("errFiles")}</p>`;
  }
}

function renderFiles(files) {
  if (!filesSection) return;
  if (!files || files.length === 0) {
    filesSection.innerHTML = `<p>${t("noFiles")}</p>`;
    renderStats();
    return;
  }

  const locale = getLocale();
  let html = '<div class="files-list">';
  files.forEach(file => {
    const modifiedDate = new Date(file.lastModifiedDateTime).toLocaleString(locale);
    const size = formatFileSize(file.size);
    const icon = getFileIcon(file.name);
    html += `
      <div class="file-item">
        <div class="file-icon">${icon}</div>
        <div class="file-info">
          <div class="file-name">${file.name}</div>
          <div class="file-details">${size} • ${t("modified")}: ${modifiedDate}</div>
        </div>
      </div>
    `;
  });
  html += "</div>";
  filesSection.innerHTML = html;

  renderStats();
}

// -------------------- Demo mode --------------------
function startDemo() {
  demoMode = true;
  showContent();

  if (demoBanner) {
    demoBanner.style.display = "block";
    demoBanner.textContent = t("demoBanner");
  }

  const now = new Date();
  const plusHours = (h) => new Date(now.getTime() + h * 60 * 60 * 1000).toISOString();

  state.profile = {
    displayName: t("demoUserName"),
    mail: "demo.user@example.com",
    userPrincipalName: "demo.user@example.com",
    jobTitle: t("demoJobTitle"),
    officeLocation: t("demoLocation")
  };

  state.emails = [
    { subject: t("demoMailSubject1"), from: { emailAddress: { name: "Contoso HR" } }, receivedDateTime: plusHours(-2), isRead: false },
    { subject: t("demoMailSubject2"), from: { emailAddress: { name: "Microsoft 365" } }, receivedDateTime: plusHours(-6), isRead: true },
    { subject: t("demoMailSubject3"), from: { emailAddress: { name: "Team Lead" } }, receivedDateTime: plusHours(-20), isRead: true }
  ];

  state.events = [
    { subject: t("demoEvent1"), start: { dateTime: plusHours(6) }, end: { dateTime: plusHours(7) }, location: { displayName: t("demoEventLoc1") } },
    { subject: t("demoEvent2"), start: { dateTime: plusHours(30) }, end: { dateTime: plusHours(31) }, location: { displayName: t("demoEventLoc2") } }
  ];

  state.files = [
    { name: "CV_Vidal_Renao.pdf", lastModifiedDateTime: plusHours(-12), size: 352000 },
    { name: "Azure-Arc-Lab-Notes.docx", lastModifiedDateTime: plusHours(-28), size: 118000 },
    { name: "Network-Diagram.png", lastModifiedDateTime: plusHours(-40), size: 890000 }
  ];

  state._lastUpdated = new Date().toISOString();

  applyTranslations();
  rerenderAll();
  renderStats();
}

// -------------------- Utilities --------------------
function formatFileSize(bytes) {
  if (!bytes || bytes === 0) return "0 Bytes";
  const k = 1024;
  const sizes = ["Bytes", "KB", "MB", "GB", "TB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + " " + sizes[i];
}

function getFileIcon(filename) {
  const ext = String(filename || "").split(".").pop().toLowerCase();
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

} // end IS_MSAL_POPUP guard
