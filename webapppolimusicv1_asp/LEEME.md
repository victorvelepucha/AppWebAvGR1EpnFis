# PoliMusic — Aplicación web tipo Spotify en ASP clásico

Clon de Spotify desarrollado en **ASP clásico (VBScript)** sobre **SQL Server**, con
identidad visual **azul** en lugar del verde de Spotify.

---

## 1. Qué se construyó

| Requisito | Estado | Dónde |
|---|---|---|
| Clon de Spotify con color azul | Hecho | `PoliMusic/css/style.css` (tema oscuro + azul `#2E7DFF`) |
| Pantalla de login | Hecho | `PoliMusic/login.asp` |
| Pantalla de registro (enrolamiento) | Hecho | `PoliMusic/register.asp` |
| Pantalla de logout | Hecho | `PoliMusic/logout.asp` |
| Pantalla principal | Hecho | `PoliMusic/home.asp` |
| CRUD de Usuarios, Canciones, Artistas, Álbumes y Playlists solo para administrador | Hecho | `PoliMusic/admin/*.asp` (motor CRUD genérico) |
| Usuario normal: reproducir canciones, consultar álbumes y artistas | Hecho | `songs.asp`, `albums.asp`, `artists.asp`, `album.asp`, `artist.asp`, `song.asp` |
| Usuario normal: consultar, crear, actualizar y eliminar **sus** playlists | Hecho | `PoliMusic/library.asp` y `PoliMusic/playlist.asp` |
| Preparado para desplegar en IIS (`C:\inetpub\wwwroot`) | Hecho | `PoliMusic/desplegar.ps1` + `PoliMusic/web.config` |

---

## 2. Cuentas de prueba

| Rol | Usuario | Contraseña | `USER_TYPE` |
|---|---|---|---|
| Administrador | `userAdmin` | `12345678` | 1 |
| Usuario normal | `epn.fis` | `12345678` | 0 |

> También existen `userNormal` / `12345678` y `username1` / `12345678`
> (`username1` guarda su contraseña como SHA‑1, igual que las cuentas nuevas).

---

## 3. Estructura del proyecto

```
webapppolimusicv1_asp/
├── BDD_PoliMusic.sql                  Script original de la base de datos
├── SongFilesAndImages_resources/      Recursos entregados (mp3 e imágenes)
├── generar_recursos.py                Genera carátulas y avatares azules
├── asp_test.py                        Arnés que ejecuta ASP clásico sin IIS
├── run_tests.py                       64 pruebas de todas las páginas
├── run_tests_crud.py                  31 pruebas de escritura (CRUD real)
├── generar_vista_previa.py            Renderiza las pantallas a HTML estático
│
└── PoliMusic/                         ← APLICACIÓN (esto es lo que se despliega)
    ├── default.asp                    Punto de entrada (redirige)
    ├── login.asp / register.asp / logout.asp
    ├── home.asp                       Pantalla principal
    ├── search.asp                     Búsqueda global
    ├── songs.asp / albums.asp / artists.asp
    ├── album.asp / artist.asp / song.asp
    ├── library.asp / playlist.asp     Playlists (CRUD del usuario normal)
    ├── profile.asp                    Perfil propio y cambio de contraseña
    ├── error404.asp
    ├── admin/
    │   ├── default.asp                Panel con indicadores
    │   ├── list.asp                   Listado genérico (buscar, ordenar, paginar)
    │   ├── edit.asp                   Formulario genérico de alta / edición
    │   └── action.asp                 Guarda y elimina (con validaciones)
    ├── api/
    │   ├── play.asp                   Incrementa el contador de reproducciones
    │   ├── playlists.asp              Playlists del usuario (JSON, para el modal)
    │   └── playlist_song.asp          Agrega una canción a una playlist (JSON)
    ├── includes/
    │   ├── common.asp                 Arranque común (incluye todo lo demás)
    │   ├── config.asp                 Configuración y cadena de conexión
    │   ├── db.asp                     Capa ADO con consultas parametrizadas
    │   ├── util.asp                   Utilidades, validaciones y escapado
    │   ├── auth.asp                   Login, sesión y control de roles
    │   ├── crud.asp                   Metadatos de las 5 entidades
    │   ├── ui.asp                     Componentes de interfaz reutilizables
    │   ├── icons.asp                  Iconos SVG en línea (sin dependencias)
    │   └── layout.asp                 Cabecera, barra lateral y reproductor
    ├── css/style.css                  Tema azul completo
    ├── js/app.js                      Reproductor, cola, modales y AJAX
    ├── assets/                        Imágenes, carátulas y archivos de audio
    ├── web.config                     Configuración de IIS (ASP, MIME, seguridad)
    ├── desplegar.ps1                  Script de despliegue (se ejecuta como admin)
    └── desplegar.bat                  Lanzador con doble clic
```

---

## 4. Despliegue en IIS — **requiere administrador**

> **Importante:** en este equipo la característica **ASP clásico (IIS‑ASP) no está
> instalada** y la cuenta `WebAvanzadas` **no pertenece al grupo Administradores**,
> por lo que el despliegue automático no se pudo ejecutar desde aquí.
> Los pasos siguientes son la única parte que falta y deben hacerse con una cuenta
> de administrador.

### Opción A — automática (recomendada)

1. Abre la carpeta `PoliMusic`.
2. Haz **clic derecho** en `desplegar.bat` → **Ejecutar como administrador**.
   (También sirve: `powershell -ExecutionPolicy Bypass -File .\desplegar.ps1` desde
   una consola elevada.)
3. Acepta el aviso de UAC y espera. El script:
   - habilita las características de IIS, incluida **IIS‑ASP**;
   - crea el grupo de aplicaciones `PoliMusicPool` (sin código administrado, 64 bits);
   - copia la aplicación a `C:\inetpub\wwwroot\PoliMusic`;
   - ajusta los permisos NTFS para `IIS AppPool\PoliMusicPool` e `IIS_IUSRS`;
   - registra la aplicación y la página 404 personalizada;
   - reinicia IIS y **verifica con peticiones HTTP reales**.

Al terminar muestra la URL y el resultado de cada verificación.

### Opción B — manual

```powershell
# 1) Habilitar IIS y ASP clásico (consola elevada)
dism /online /enable-feature /featurename:IIS-WebServerRole /all
dism /online /enable-feature /featurename:IIS-ASP /all
dism /online /enable-feature /featurename:IIS-DefaultDocument /all
dism /online /enable-feature /featurename:IIS-RequestFiltering /all

# 2) Copiar la aplicación
robocopy "<ruta>\PoliMusic" "C:\inetpub\wwwroot\PoliMusic" /MIR

# 3) Permisos de lectura
icacls "C:\inetpub\wwwroot\PoliMusic" /grant "IIS_IUSRS:(OI)(CI)(RX)" /T

# 4) Reiniciar IIS
iisreset
```

### Variantes del script

| Comando | Efecto |
|---|---|
| `desplegar.bat` | Publica en `http://localhost/PoliMusic/` (recomendado) |
| `desplegar.bat -EnLaRaiz` | Publica en `http://localhost/` (raíz del sitio) |
| `desplegar.bat -Debug` | Activa `scriptErrorSentToBrowser` para depurar |
| `desplegar.bat -SaltarCaracteristicas` | No toca las características de Windows |

---

## 5. Vista previa sin IIS

Como ASP clásico solo se ejecuta sobre IIS, se incluye una **vista previa estática**
generada con el propio código de la aplicación (`PoliMusic/_preview/`). Permite
revisar el diseño y la navegación sin desplegar nada:

```
python -m http.server 8099 --bind 127.0.0.1
# abrir http://127.0.0.1:8099/PoliMusic/_preview/index.html
```

Los formularios y la reproducción de audio requieren IIS con ASP clásico.

---

## 6. Cómo está resuelto el acceso a datos

- **Consultas 100 % parametrizadas** (`includes/db.asp`). No se concatena nunca
  entrada del usuario en el SQL → sin inyección SQL.
- **Proveedor OLE DB con respaldo automático**: se intentan `MSOLEDBSQL`,
  `MSOLEDBSQL19`, `SQLNCLI11`, `SQLNCLI10` y `SQLOLEDB`, y el que funciona se
  memoriza en `Application` para no reintentarlo.
- **Contraseñas con SHA‑1 calculado en SQL Server** (`HASHBYTES`), de modo que el
  formato coincide con el que ya usaba la base de datos. El login acepta además
  las contraseñas heredadas en texto plano de la carga inicial.
- **Salida escapada**: todo dato proveniente del usuario pasa por `H()` antes de
  escribirse en HTML → sin XSS.
- **Codificación**: las páginas responden en UTF‑8 y los archivos `.asp` se
  mantienen en ASCII puro (los acentos se escriben con entidades HTML), lo que
  evita los problemas clásicos de codificación de ASP clásico.

## 7. Reglas de negocio implementadas

| Regla | Implementación |
|---|---|
| Solo el administrador hace CRUD de usuarios, canciones, artistas, álbumes y playlists | `RequireAdmin` + `admin/action.asp`; el usuario normal recibe **403** |
| El usuario normal reproduce, consulta álbumes y artistas | Todas las pantallas salvo `admin/` |
| El usuario normal hace CRUD de **sus** playlists | `CanManagePlaylist()` en `includes/auth.asp` |
| Playlists públicas / privadas | `TBL_PLAYLIST.TYPE` (0 = pública, 1 = privada) según `TBL_CATALOG`; las privadas ajenas devuelven 403 |
| Borrado con integridad referencial | `EntityDelete()` limpia `TBL_TRACKLIST`, `TBL_PLAYER` y `TBL_PLAYLIST_SONG` antes de borrar |
| El administrador no puede eliminarse a sí mismo ni quitarse el rol | Validado en `EntityDelete()` y en `admin/action.asp` |

---

## 8. Verificación realizada

Se construyó un **arnés que ejecuta el ASP clásico real fuera de IIS**
(`asp_test.py`: resuelve los `#include`, traduce `<% %>` a VBScript y simula los
objetos `Request`, `Response`, `Server`, `Session` y `Application`, conectándose
de verdad a SQL Server con ADO).

| Batería | Pruebas | Resultado |
|---|---|---|
| `run_tests.py` — todas las páginas con sesión anónima, normal y administrador | 64 | **64 correctas** |
| `run_tests_crud.py` — altas, modificaciones, bajas, validaciones y permisos contra la base real | 31 | **31 correctas** |

Entre otras cosas se comprobó que: las contraseñas nuevas se guardan como SHA‑1, la
contraseña no se pierde al editar un usuario sin escribirla, los duplicados y los
correos inválidos se rechazan, el borrado en cascada deja la base consistente, un
usuario normal **no** puede modificar playlists ajenas, y la base de datos queda
exactamente en su estado original al terminar las pruebas.

### Incidencias reales detectadas y corregidas por el arnés

1. `to` es palabra reservada en VBScript y se usaba como variable en la paginación.
2. Un comentario de documentación contenía `<% %>`, que ASP habría **ejecutado**.
3. `IIf` en VBScript evalúa **las dos ramas**: se leía una columna inexistente.
4. El formulario de "quitar canción" enviaba `playlistId` mientras el código leía `id`.
5. `FmtDate` estaba declarada con dos parámetros y se llamaba con uno.
6. `CONVERT(..., 23)` de SQL Server falla en esta versión del motor; se sustituyó por
   `TRY_CONVERT(..., 120)`.

---

## 9. Notas sobre los recursos

- Solo se entregó **un** archivo de audio (`bensound-adventure.mp3`) para las cinco
  canciones del catálogo; se replicó con los nombres que referencia la base para
  que las cinco pistas sean reproducibles.
- La base referencia imágenes de álbumes y artistas que no venían en la carpeta de
  recursos, así que se **generaron** con la identidad azul de la aplicación
  (`generar_recursos.py`, con Pillow): carátulas tipo vinilo y avatares con iniciales.
- Si un archivo de audio no existe, la fila se marca como *"Archivo no disponible"*
  en lugar de fallar.
