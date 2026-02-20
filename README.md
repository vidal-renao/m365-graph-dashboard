# 🚀 My Microsoft 365 Dashboard

Aplicación web que se conecta a Microsoft Graph API para mostrar tu información de Microsoft 365.

## 📋 ¿Qué hace esta aplicación?

- ✅ Muestra tu perfil de Microsoft
- ✅ Lista tus últimos 5 emails
- ✅ Muestra tus próximos eventos de calendario
- ✅ Lista tus archivos recientes de OneDrive

## 🔧 Configuración completada

La aplicación ya está configurada con tu Application ID de Azure:
- **Application ID:** `58d4f2d3-5598-401e-a2ff-a01806d304e7`
- **Redirect URI:** `http://localhost:8080`

## 🚀 Cómo usar la aplicación

### Opción 1: Usar el servidor Python incluido (Recomendado)

1. Abre una terminal en la carpeta del proyecto
2. Ejecuta uno de estos comandos según tu versión de Python:

**Python 3:**
```bash
python -m http.server 8080
```

**Python 2:**
```bash
python -m SimpleHTTPServer 8080
```

3. Abre tu navegador y ve a: `http://localhost:8080`
4. Click en "Iniciar Sesión con Microsoft"
5. ¡Listo! Verás tus datos de Microsoft 365

### Opción 2: Usar Node.js (si tienes npm instalado)

1. Instala un servidor HTTP simple:
```bash
npm install -g http-server
```

2. En la carpeta del proyecto, ejecuta:
```bash
http-server -p 8080
```

3. Abre `http://localhost:8080` en tu navegador

### Opción 3: Usar extensión de VS Code

Si usas Visual Studio Code:
1. Instala la extensión "Live Server"
2. Click derecho en `index.html`
3. Selecciona "Open with Live Server"
4. **IMPORTANTE:** Cambia el puerto a 8080 en la configuración

## ⚠️ Importante

- **DEBES usar `http://localhost:8080`** exactamente (no otro puerto, no 127.0.0.1)
- La primera vez que inicies sesión, Microsoft te pedirá permisos para acceder a tus datos
- Usa tu cuenta personal de Microsoft (@outlook.com, @hotmail.com, etc.) o tu cuenta corporativa

## 🧪 Probar la aplicación

1. Inicia la aplicación en `http://localhost:8080`
2. Click en "Iniciar Sesión con Microsoft"
3. Se abrirá una ventana popup de Microsoft
4. Ingresa tus credenciales de Microsoft
5. Acepta los permisos solicitados
6. ¡La aplicación cargará automáticamente tus datos!

## 📝 Notas

- Si ves errores de CORS, asegúrate de estar usando `localhost:8080` (no otra dirección)
- Algunos datos pueden no aparecer si no tienes configurado el servicio (ej: OneDrive, Exchange)
- La aplicación funciona con cuentas personales y corporativas de Microsoft

## 🔒 Seguridad

- Esta aplicación NO almacena tus credenciales
- Usa autenticación OAuth2 de Microsoft (MSAL.js)
- Los tokens se guardan en localStorage de tu navegador
- Puedes cerrar sesión en cualquier momento

## 📚 Tecnologías usadas

- HTML5
- CSS3
- JavaScript (ES6+)
- MSAL.js 2.0 (Microsoft Authentication Library)
- Microsoft Graph API v1.0

## 🐛 Solución de problemas

**Error: "Redirect URI mismatch"**
- Verifica que estés usando `http://localhost:8080` exactamente
- Verifica la configuración en Azure Portal

**Error: "CORS"**
- Usa un servidor HTTP local (no abras el archivo directamente)
- Usa el puerto 8080 configurado

**No aparecen emails/calendario/archivos**
- Verifica que tengas esos servicios configurados en tu cuenta
- Algunos servicios solo están disponibles en cuentas corporativas

## 📞 ¿Necesitas ayuda?

Este proyecto fue creado para demostrar el uso de Microsoft Graph API y calificar para el Microsoft 365 Developer Program.

---

**¡Disfruta explorando tu dashboard de Microsoft 365!** 🎉
