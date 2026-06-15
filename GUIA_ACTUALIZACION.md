# GUIA DE ACTUALIZACION - APP DISPONIBLES
URL publica: https://danielvega-inmobiliaria.github.io/Disponibles/

---

## ARCHIVOS DE LA CARPETA

| Archivo | Funcion |
|---|---|
| MENU.bat | Menu principal — ejecutar esto para todo |
| generar_html.py | Genera la app HTML con imagenes |
| actualizar_disponibles.py | Actualiza precios desde listas M2 y Obring |
| disponibles_mobile.html | App generada (se sube a GitHub) |
| index.html | Redireccion para GitHub Pages |
| IMAGENES/ | Carpeta de fotos de edificios |

---

## CASO 1: ACTUALIZAR PRECIOS

Cuando llega una nueva lista de precios de M2 o Obring:

1. Reemplazar el archivo Excel en su carpeta correspondiente:
   - M2: `DESARROLLOS/M2 DESARROLLOS/LISTAS DE PRECIO/`
   - Obring: `DESARROLLOS/OBRING/`

2. Abrir **MENU.bat**

3. Presionar **1** → Actualizar precios
   - Se genera `DISPONIBLES_actualizado_FECHA.xlsx` en esta carpeta

4. Presionar **2** → Generar HTML
   - La app toma automaticamente el archivo actualizado mas reciente
   - Verificar que diga "Fachadas embebidas: 12" (o mas)

5. Abrir Git Bash en esta carpeta y ejecutar:
   ```
   git add disponibles_mobile.html index.html
   git commit -m "Actualizar precios"
   git push
   ```

6. Esperar 1-2 minutos y verificar en el celular.

---

## CASO 2: AGREGAR IMAGENES DE UN EDIFICIO NUEVO

1. Nombrar las fotos con este formato exacto:
   ```
   NOMBRE EDIFICIO - TIPO.jpg
   ```
   Ejemplos:
   ```
   MORENO 323 - FACHADA.jpg       <- portada de la tarjeta (obligatoria)
   MORENO 323 - PILETA.jpg        <- galeria interior
   MORENO 323 - QUINCHO.jpg       <- galeria interior
   MORENO 323 - DEPTO.jpg         <- galeria interior
   ```
   - El nombre antes del " - " debe coincidir exactamente con la direccion en DISPONIBLES.xlsx
   - La FACHADA aparece como miniatura en la tarjeta y como imagen principal en el detalle
   - Las otras imagenes (hasta 4) aparecen en la grilla del detalle
   - Formatos validos: jpg, jpeg, png, tif, webp

2. Copiar los archivos a la carpeta `IMAGENES/`

3. Abrir **MENU.bat** → presionar **2** → Generar HTML

4. Verificar que aparezca "FACHADA  Moreno 323" (sin "??") en el output

5. Subir a GitHub:
   ```
   git add disponibles_mobile.html index.html
   git commit -m "Agregar imagenes Moreno 323"
   git push
   ```

---

## CASO 3: ACTUALIZAR DISPONIBLES.xlsx MANUALMENTE

Si se agregan o quitan propiedades directamente en el Excel:

1. Editar `DISPONIBLES.xlsx` en la carpeta DESARROLLOS (un nivel arriba)

2. Abrir **MENU.bat** → presionar **2** → Generar HTML

3. Subir a GitHub (ver paso 5 del Caso 1)

---

## SUBIR A GITHUB (resumen)

Abrir Git Bash en esta carpeta (click derecho → "Git Bash Here") y ejecutar:

```bash
git add disponibles_mobile.html index.html
git commit -m "descripcion del cambio"
git push
```

La URL publica se actualiza en 1-2 minutos automaticamente.

---

## PROBAR EN RED LOCAL (sin internet)

1. Abrir **MENU.bat** → presionar **4**
2. Aparece la URL local, por ejemplo: `http://192.168.1.5:8080/disponibles_mobile.html`
3. Abrir esa URL en el celular (mismo WiFi)
4. Presionar Ctrl+C en la terminal para detener

---

## SOLUCION DE PROBLEMAS

**"No module named openpyxl"** → MENU.bat → opcion 0 (instalar dependencias)

**Las imagenes no aparecen** → Verificar que el nombre del archivo coincida con la columna UBICACION del Excel. Ejecutar opcion 2 y ver si aparece "??" antes del nombre del edificio.

**El push de Git falla** → Ejecutar `git pull origin main` primero, luego volver a hacer push.

**La app no se actualiza en el celular** → Esperar 2 minutos. Si sigue igual, limpiar cache del browser.
