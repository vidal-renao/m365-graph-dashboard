# 🚀 Microsoft 365 Dashboard (Microsoft Graph + MSAL)

![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-API-blue)
![MSAL.js](https://img.shields.io/badge/MSAL.js-Auth-orange)
![JavaScript](https://img.shields.io/badge/JavaScript-ES6%2B-yellow)
![OAuth](https://img.shields.io/badge/OAuth%202.0-Authorization-green)
![Languages](https://img.shields.io/badge/Languages-ES%20%7C%20EN%20%7C%20DE-purple)

**Languages / Idiomas / Sprachen:**  
[🇪🇸 Español](#-español) | [🇬🇧 English](#-english) | [🇩🇪 Deutsch](#-deutsch)

---

## 🇪🇸 Español

### 📋 Descripción
Dashboard moderno que muestra datos de **Microsoft 365** usando **Microsoft Graph API** y autenticación con **MSAL.js**.

Incluye **modo Demo (sin login)** para que cualquiera pueda ver el dashboard en GitHub sin configurar Microsoft 365.

### ✨ Funcionalidades
- 🔐 Login con Microsoft (MSAL.js / OAuth 2.0)
- 👤 Perfil de usuario
- 📧 Emails recientes (leídos / no leídos)
- 📅 Próximos eventos (7 días)
- 📁 Archivos recientes (OneDrive)
- 🌍 UI multi-idioma (ES / EN / DE)
- 👀 **Modo Demo (sin login)**

### 📸 Capturas
Crea una carpeta `screenshots/` y añade estas imágenes (nombres sugeridos):
- `screenshots/01-login.png`
- `screenshots/02-dashboard.png`
- `screenshots/03-language-switch.png`

Luego puedes incluirlas así:
```md
![Login](screenshots/01-login.png)
![Dashboard](screenshots/02-dashboard.png)
![Languages](screenshots/03-language-switch.png)
```

### 🛠️ Tecnologías
- HTML5 / CSS3 / JavaScript (Vanilla)
- Microsoft Graph API v1.0
- MSAL.js 2.x

### 🔑 Permisos Graph
- `User.Read`
- `Mail.Read`
- `Calendars.Read`
- `Files.Read.All`

### 🚀 Ejecutar en local
```bash
python -m http.server 8080
```
Abrir: `http://localhost:8080`

### 👀 Modo Demo (sin login)
En la página, pulsa **“Ver demo (sin login)”**.  
Para ver datos reales: inicia sesión con Microsoft.

### 🧠 What I learned / Lo que demuestra
- Integración con Microsoft Graph
- Autenticación SPA con MSAL.js
- Permisos/scopes y manejo de tokens
- Renderizado y estado en frontend sin frameworks
- i18n simple y mantenible (ES/EN/DE)

---

## 🇬🇧 English

### 📋 Overview
A modern dashboard that displays **Microsoft 365** data using **Microsoft Graph API** and **MSAL.js** authentication.

Includes a **Demo mode (no login)** so anyone can preview the UI directly from GitHub.

### ✨ Features
- Microsoft sign-in (MSAL.js / OAuth 2.0)
- User profile
- Recent emails
- Upcoming events (7 days)
- Recent OneDrive files
- Multi-language UI (ES / EN / DE)
- **Demo mode (no login)**

### 🚀 Run locally
```bash
python -m http.server 8080
```
Open: `http://localhost:8080`

### 👀 Demo Mode
Click **“View demo (no login)”**.  
Sign in to see your real Microsoft 365 data.

---

## 🇩🇪 Deutsch

### 📋 Beschreibung
Modernes Dashboard für **Microsoft 365** mit **Microsoft Graph API** und Authentifizierung über **MSAL.js**.

Enthält einen **Demo-Modus (ohne Login)**, damit man die UI sofort testen kann.

### ✨ Funktionen
- Microsoft-Anmeldung (MSAL.js / OAuth 2.0)
- Benutzerprofil
- Letzte E-Mails
- Nächste Termine (7 Tage)
- Letzte OneDrive-Dateien
- Mehrsprachige UI (ES / EN / DE)
- **Demo-Modus (ohne Login)**

### 🚀 Lokal starten
```bash
python -m http.server 8080
```
Öffnen: `http://localhost:8080`

### 👀 Demo-Modus
Klicke **„Demo ansehen (ohne Login)“**.  
Für echte Daten: mit Microsoft anmelden.

---

## 👤 Author
**Vidal Reñao Lopelo**  
LinkedIn: https://www.linkedin.com/in/vidalrenao

⭐ If you find this useful, please star the repo!
