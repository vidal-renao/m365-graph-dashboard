# 🚀 Microsoft 365 Dashboard

> A modern, real-time dashboard connecting to **Microsoft Graph API** — built with vanilla JavaScript and MSAL.js authentication.

![Dashboard Preview](screenshots/dashboard-main.png)

---

## 🌍 Available Languages / Idiomas disponibles / Verfügbare Sprachen

| 🇪🇸 [Español](#-español) | 🇬🇧 [English](#-english) | 🇩🇪 [Deutsch](#-deutsch) |
|:---:|:---:|:---:|

---

## 🇪🇸 Español

### 📋 Descripción

Aplicación web moderna que muestra datos de **Microsoft 365** en tiempo real utilizando la **Microsoft Graph API**. Desarrollada con **JavaScript puro** y **MSAL.js** para autenticación segura mediante OAuth 2.0.

### ✨ Funcionalidades

| Función | Descripción |
|---|---|
| 🔐 Autenticación | OAuth 2.0 con MSAL.js |
| 👤 Perfil de usuario | Nombre, email, puesto y ubicación |
| 📊 Quick Stats | Emails no leídos, próximo evento, archivos recientes |
| 📧 Últimos emails | Bandeja de entrada en tiempo real |
| 📅 Calendario | Próximos eventos de los siguientes 7 días |
| 📁 OneDrive | Archivos recientes |
| 🌍 Multi-idioma | Interfaz en ES / EN / DE |
| 🎨 UI moderna | Diseño responsive y accesible |

### 📸 Capturas de pantalla

| Login | Dashboard |
|---|---|
| ![Login](screenshots/login.png) | ![Dashboard](screenshots/dashboard-main.png) |

| Perfil | Quick Stats |
|---|---|
| ![Perfil](screenshots/profile.png) | ![Stats](screenshots/quick-stats.png) |

| Emails | Calendario |
|---|---|
| ![Emails](screenshots/emails.png) | ![Calendario](screenshots/calendar.png) |

| Archivos | Cambio de idioma |
|---|---|
| ![Archivos](screenshots/files.png) | ![Idiomas](screenshots/Switching_language.png) |

### 🚀 Cómo ejecutarlo

```bash
# Clona el repositorio
git clone https://github.com/vidal-renao/m365-graph-dashboard.git
cd m365-graph-dashboard

# Inicia el servidor local
python -m http.server 8080

# Abre en el navegador
http://localhost:8080
```

### ⚙️ Requisitos

- Cuenta de Microsoft 365
- Registro de aplicación en **Azure Active Directory**
- Permisos de Microsoft Graph: `User.Read`, `Mail.Read`, `Calendars.Read`, `Files.Read`

---

## 🇬🇧 English

### 📋 Overview

A modern web application displaying **Microsoft 365** data in real time using the **Microsoft Graph API**. Built with **vanilla JavaScript** and **MSAL.js** for secure OAuth 2.0 authentication.

### ✨ Features

| Feature | Description |
|---|---|
| 🔐 Authentication | OAuth 2.0 with MSAL.js |
| 👤 User profile | Name, email, job title, location |
| 📊 Quick Stats | Unread emails, next event, recent files |
| 📧 Latest emails | Real-time inbox |
| 📅 Calendar | Upcoming events for the next 7 days |
| 📁 OneDrive | Recent files |
| 🌍 Multi-language | UI in ES / EN / DE |
| 🎨 Modern UI | Responsive and accessible design |

### 📸 Screenshots

| Login | Dashboard |
|---|---|
| ![Login](screenshots/login.png) | ![Dashboard](screenshots/dashboard-main.png) |

| Profile | Quick Stats |
|---|---|
| ![Profile](screenshots/profile.png) | ![Stats](screenshots/quick-stats.png) |

| Emails | Calendar |
|---|---|
| ![Emails](screenshots/emails.png) | ![Calendar](screenshots/calendar.png) |

| Files | Language switch |
|---|---|
| ![Files](screenshots/files.png) | ![Language](screenshots/Switching_language.png) |

### 🚀 How to run

```bash
# Clone the repository
git clone https://github.com/vidal-renao/m365-graph-dashboard.git
cd m365-graph-dashboard

# Start local server
python -m http.server 8080

# Open in browser
http://localhost:8080
```

### ⚙️ Requirements

- Microsoft 365 account
- App registration in **Azure Active Directory**
- Microsoft Graph permissions: `User.Read`, `Mail.Read`, `Calendars.Read`, `Files.Read`

---

## 🇩🇪 Deutsch

### 📋 Übersicht

Eine moderne Webanwendung, die **Microsoft 365**-Daten in Echtzeit über die **Microsoft Graph API** anzeigt. Entwickelt mit **Vanilla JavaScript** und **MSAL.js** für sichere OAuth 2.0-Authentifizierung.

### ✨ Funktionen

| Funktion | Beschreibung |
|---|---|
| 🔐 Authentifizierung | OAuth 2.0 mit MSAL.js |
| 👤 Benutzerprofil | Name, E-Mail, Stelle, Standort |
| 📊 Quick Stats | Ungelesene E-Mails, nächster Termin, aktuelle Dateien |
| 📧 Letzte E-Mails | Posteingang in Echtzeit |
| 📅 Kalender | Nächste Termine der folgenden 7 Tage |
| 📁 OneDrive | Zuletzt verwendete Dateien |
| 🌍 Mehrsprachig | Oberfläche in ES / EN / DE |
| 🎨 Modernes UI | Responsives und barrierefreies Design |

### 📸 Screenshots

| Login | Dashboard |
|---|---|
| ![Login](screenshots/login.png) | ![Dashboard](screenshots/dashboard-main.png) |

| Profil | Quick Stats |
|---|---|
| ![Profil](screenshots/profile.png) | ![Stats](screenshots/quick-stats.png) |

| E-Mails | Kalender |
|---|---|
| ![E-Mails](screenshots/emails.png) | ![Kalender](screenshots/calendar.png) |

| Dateien | Sprachwechsel |
|---|---|
| ![Dateien](screenshots/files.png) | ![Sprache](screenshots/Switching_language.png) |

### 🚀 Lokal starten

```bash
# Repository klonen
git clone https://github.com/vidal-renao/m365-graph-dashboard.git
cd m365-graph-dashboard

# Lokalen Server starten
python -m http.server 8080

# Im Browser öffnen
http://localhost:8080
```

### ⚙️ Voraussetzungen

- Microsoft 365-Konto
- App-Registrierung in **Azure Active Directory**
- Microsoft Graph-Berechtigungen: `User.Read`, `Mail.Read`, `Calendars.Read`, `Files.Read`

---

## 🛠️ Tech Stack

![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![Microsoft](https://img.shields.io/badge/Microsoft_Graph-0078D4?style=for-the-badge&logo=microsoft&logoColor=white)
![MSAL](https://img.shields.io/badge/MSAL.js_2.0-00A4EF?style=for-the-badge&logo=microsoft-azure&logoColor=white)
![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![CSS3](https://img.shields.io/badge/CSS3-1572B6?style=for-the-badge&logo=css3&logoColor=white)

---

## 👤 Author

**Vidal Reñao Lopelo**

[![GitHub](https://img.shields.io/badge/GitHub-vidal--renao-181717?style=for-the-badge&logo=github)](https://github.com/vidal-renao)

---

⭐ Si te gusta este proyecto, ¡dale una estrella! / If you like this project, give it a star! / Wenn Ihnen dieses Projekt gefällt, geben Sie ihm einen Stern!
