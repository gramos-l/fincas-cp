# Fincas C&P

Mapa interactivo en 3D de fincas de boda a 20-24 km de Calle de Ayala 71 (Madrid), con fichas por finca.

## Qué hace

- Mapa que, al acercarte (zoom 15+), levanta los edificios con su altura real (OpenStreetMap).
- Chinchetas de las fincas (rojas = pendientes, verdes = visitadas en persona) con **agrupación** al alejar el zoom.
- Ficha al pinchar una chincheta: precio, aforo, catering, amenidades, valoraciones (Google / bodas.net), teléfono, enlaces a Google Maps, ruta desde Ayala 71, web e Instagram.
- Filtros: Todas / Dentro de 20 km / 20-24 km / Otros espacios. Lista con buscador.
- Edición local (botón **Editar** en la ficha): marcar visitada, notas propias, mover la chincheta — se guarda en el móvil y se puede exportar/importar.
- **Cargar mi Excel** (en Opciones): si prefieres rellenar un Excel con tus propios datos en vez de editar en la app, lo cargas ahí — se guarda en este dispositivo y el mapa lo usa en cada visita (botón **Quitar mi Excel** para volver al del repositorio).
- **Exportar a Excel** (en Opciones): descarga un `.xlsx` con las mismas hojas y columnas que el original, más una columna de Estado con tus ediciones de la app.
- **Sincronizar con otra persona** (en Opciones): en tiempo real, vía Firebase — ver sección de abajo. Si no se configura, cada dispositivo sigue guardando sus propios cambios.
- Intro: el mapa arranca viendo la Tierra desde el espacio y baja hasta Ayala 71.
- Caché en el móvil (service worker): las visitas siguientes cargan más rápido.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `index.html` | Estructura de la página. |
| `styles.css` | Diseño y colores (bloque de variables al principio). |
| `app.js` | Mapa, chinchetas, fichas, lista, filtros y lectura del Excel. |
| `editor.js` | Edición local de cada finca (visitada, notas, teléfono…). |
| `sw.js` | Service worker: caché de teselas y librería. |
| `fincas.json` | Copia de los datos combinando las hojas "Fincas 20 km", "Justo fuera 20-24 km" y "Otros espacios". |
| `Fincas_bodas_20km_Ayala71.xlsx` | El Excel original. Si subes una versión nueva con este mismo nombre, el mapa la lee en vez del JSON. |

## Cómo actualizar los datos

1. Edita el Excel (hojas **Fincas 20 km**, **Justo fuera 20-24 km**, **Otros espacios**, con la cabecera en la fila 1).
2. Sustituye `Fincas_bodas_20km_Ayala71.xlsx` por la nueva versión (mismo nombre exacto).
3. Si lo despliegas (GitHub Pages u otro), recarga con `?v=2` al final o dos veces para saltarte cualquier copia guardada por el navegador.

Si una fila no trae Latitud/Longitud, el mapa busca la dirección en Nominatim (OpenStreetMap) una sola vez y guarda el resultado en el móvil.

## Activar la sincronización entre dispositivos

Desactivada por defecto (`FIREBASE_CONFIG` vacío en `editor.js`, sección "sincronizar entre
dispositivos"). Para activarla:

1. Crea un proyecto gratuito en [console.firebase.google.com](https://console.firebase.google.com).
2. **Build → Firestore Database → Crear base de datos** (modo producción, cualquier región).
3. **Build → Authentication → Sign-in method → Anonymous** → activar.
4. En las reglas de Firestore (pestaña **Rules**), pega:
   ```
   rules_version = '2';
   service cloud.firestore {
     match /databases/{database}/documents {
       match /boards/{boardId} {
         allow read, write: if request.auth != null;
       }
     }
   }
   ```
5. **Configuración del proyecto (⚙️) → tus apps → Añadir app → Web (`</>`)**, dale un nombre y copia
   el objeto `firebaseConfig` que te muestra.
6. Pega esos valores en `FIREBASE_CONFIG` al principio de la sección de sincronización en
   `editor.js`.
7. (Recomendado) En Google Cloud Console → APIs y servicios → Credenciales, restringe la API key
   de ese proyecto para que solo funcione desde el dominio donde publiques el mapa — la clave viaja
   en el código de la página, así que cualquiera que la vea podría usarla si no la restringes.

Con eso, cualquier persona que entre con el mismo código de 8 letras (botón **Crear código para
compartir** / **Unirme con un código**, o el enlace que genera "Copiar enlace para compartir") verá
los mismos cambios en tiempo real. El plan gratuito de Firebase (Spark) sobra de sobra para esto.

## Publicar

No hay compilación. Cualquier hosting de archivos estáticos sirve (GitHub Pages: Settings → Pages → Deploy from a branch → `main` / `(root)`).

## Créditos

Mismo stack que [mapa_vermut](../mapa_vermut): MapLibre GL JS 6.15.0, teselas de [OpenFreeMap](https://openfreemap.org) (© OpenStreetMap), SheetJS para leer el Excel.
