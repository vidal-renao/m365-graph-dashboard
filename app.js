// Configuration for MSAL (Microsoft Authentication Library)
const msalConfig = {
    auth: {
        clientId: '58d4f2d3-5598-401e-a2ff-a01806d304e7',
        authority: 'https://login.microsoftonline.com/common',
        redirectUri: 'http://localhost:8080'
    },
    cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: false
    }
};

// Permissions we need
const loginRequest = {
    scopes: ['User.Read', 'Mail.Read', 'Calendars.Read', 'Files.Read.All']
};

// Initialize MSAL
const msalInstance = new msal.PublicClientApplication(msalConfig);

// DOM Elements
const loginButton = document.getElementById('login-button');
const logoutButton = document.getElementById('logout-button');
const loginSection = document.getElementById('login-section');
const content = document.getElementById('content');
const userName = document.getElementById('user-name');
const profileSection = document.getElementById('profile-section');
const emailSection = document.getElementById('email-section');
const calendarSection = document.getElementById('calendar-section');
const filesSection = document.getElementById('files-section');

// Event Listeners
loginButton.addEventListener('click', login);
logoutButton.addEventListener('click', logout);

// Check if user is already logged in
checkAccount();

// Functions
async function checkAccount() {
    const currentAccounts = msalInstance.getAllAccounts();
    if (currentAccounts.length > 0) {
        // User is already logged in
        showContent();
        loadUserData();
    }
}

async function login() {
    try {
        const loginResponse = await msalInstance.loginPopup(loginRequest);
        console.log('Login successful:', loginResponse);
        showContent();
        loadUserData();
    } catch (error) {
        console.error('Login error:', error);
        alert('Error al iniciar sesión: ' + error.message);
    }
}

function logout() {
    const currentAccounts = msalInstance.getAllAccounts();
    if (currentAccounts.length > 0) {
        msalInstance.logoutPopup({
            account: currentAccounts[0]
        });
    }
    hideContent();
}

function showContent() {
    loginSection.style.display = 'none';
    content.style.display = 'block';
}

function hideContent() {
    loginSection.style.display = 'block';
    content.style.display = 'none';
}

async function getAccessToken() {
    const currentAccounts = msalInstance.getAllAccounts();
    if (currentAccounts.length === 0) {
        throw new Error('No hay usuario autenticado');
    }

    const request = {
        scopes: loginRequest.scopes,
        account: currentAccounts[0]
    };

    try {
        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;
    } catch (error) {
        console.log('Silent token acquisition failed, acquiring token using popup');
        const response = await msalInstance.acquireTokenPopup(request);
        return response.accessToken;
    }
}

async function loadUserData() {
    try {
        const accessToken = await getAccessToken();
        
        // Load all data in parallel
        await Promise.all([
            loadProfile(accessToken),
            loadEmails(accessToken),
            loadCalendar(accessToken),
            loadFiles(accessToken)
        ]);
    } catch (error) {
        console.error('Error loading user data:', error);
        alert('Error al cargar los datos: ' + error.message);
    }
}

async function loadProfile(accessToken) {
    try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me', {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const profile = await response.json();
        
        // Display profile
        userName.textContent = profile.displayName || 'Usuario';
        profileSection.innerHTML = `
            <div class="profile-info">
                <p><strong>Nombre:</strong> ${profile.displayName || 'N/A'}</p>
                <p><strong>Email:</strong> ${profile.mail || profile.userPrincipalName || 'N/A'}</p>
                <p><strong>Puesto:</strong> ${profile.jobTitle || 'N/A'}</p>
                <p><strong>Ubicación:</strong> ${profile.officeLocation || 'N/A'}</p>
            </div>
        `;
    } catch (error) {
        console.error('Error loading profile:', error);
        profileSection.innerHTML = `<p class="error">❌ Error al cargar el perfil</p>`;
    }
}

async function loadEmails(accessToken) {
    try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me/messages?$top=5&$select=subject,from,receivedDateTime,isRead', {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        const emails = data.value;

        if (emails.length === 0) {
            emailSection.innerHTML = '<p>No hay emails recientes</p>';
            return;
        }

        let emailHTML = '<div class="email-list">';
        emails.forEach(email => {
            const date = new Date(email.receivedDateTime).toLocaleString('es-ES');
            const readClass = email.isRead ? 'read' : 'unread';
            emailHTML += `
                <div class="email-item ${readClass}">
                    <div class="email-subject">${email.subject || '(Sin asunto)'}</div>
                    <div class="email-from">De: ${email.from?.emailAddress?.name || 'Desconocido'}</div>
                    <div class="email-date">${date}</div>
                </div>
            `;
        });
        emailHTML += '</div>';

        emailSection.innerHTML = emailHTML;
    } catch (error) {
        console.error('Error loading emails:', error);
        emailSection.innerHTML = `<p class="error">❌ Error al cargar emails (puede que no tengas buzón configurado)</p>`;
    }
}

async function loadCalendar(accessToken) {
    try {
        const now = new Date().toISOString();
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${now}&endDateTime=${new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString()}&$top=5&$select=subject,start,end,location`, {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        const events = data.value;

        if (events.length === 0) {
            calendarSection.innerHTML = '<p>No hay eventos próximos en los próximos 7 días</p>';
            return;
        }

        let calendarHTML = '<div class="calendar-list">';
        events.forEach(event => {
            const startDate = new Date(event.start.dateTime).toLocaleString('es-ES');
            const endDate = new Date(event.end.dateTime).toLocaleString('es-ES');
            calendarHTML += `
                <div class="calendar-item">
                    <div class="event-subject">${event.subject || '(Sin título)'}</div>
                    <div class="event-time">📅 ${startDate}</div>
                    ${event.location?.displayName ? `<div class="event-location">📍 ${event.location.displayName}</div>` : ''}
                </div>
            `;
        });
        calendarHTML += '</div>';

        calendarSection.innerHTML = calendarHTML;
    } catch (error) {
        console.error('Error loading calendar:', error);
        calendarSection.innerHTML = `<p class="error">❌ Error al cargar calendario</p>`;
    }
}

async function loadFiles(accessToken) {
    try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me/drive/recent?$top=5', {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        const files = data.value;

        if (files.length === 0) {
            filesSection.innerHTML = '<p>No hay archivos recientes</p>';
            return;
        }

        let filesHTML = '<div class="files-list">';
        files.forEach(file => {
            const modifiedDate = new Date(file.lastModifiedDateTime).toLocaleString('es-ES');
            const size = formatFileSize(file.size);
            const icon = getFileIcon(file.name);
            filesHTML += `
                <div class="file-item">
                    <div class="file-icon">${icon}</div>
                    <div class="file-info">
                        <div class="file-name">${file.name}</div>
                        <div class="file-details">${size} • Modificado: ${modifiedDate}</div>
                    </div>
                </div>
            `;
        });
        filesHTML += '</div>';

        filesSection.innerHTML = filesHTML;
    } catch (error) {
        console.error('Error loading files:', error);
        filesSection.innerHTML = `<p class="error">❌ Error al cargar archivos (puede que no tengas OneDrive configurado)</p>`;
    }
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

function getFileIcon(filename) {
    const ext = filename.split('.').pop().toLowerCase();
    const icons = {
        'pdf': '📄',
        'doc': '📝',
        'docx': '📝',
        'xls': '📊',
        'xlsx': '📊',
        'ppt': '📊',
        'pptx': '📊',
        'txt': '📃',
        'jpg': '🖼️',
        'jpeg': '🖼️',
        'png': '🖼️',
        'gif': '🖼️',
        'zip': '📦',
        'rar': '📦'
    };
    return icons[ext] || '📄';
}
