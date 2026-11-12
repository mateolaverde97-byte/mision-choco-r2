# Misión Chocó · Ronda 2 · Propuestas conjuntas

Web para que los líderes completen las 21 fichas de la ronda 2. Se publica en GitHub Pages y guarda en una hoja de Google a través de Apps Script. Todo queda en línea y visible para todos.

## Archivos
- `index.html`: la página completa (datos, diseño y lógica en un solo archivo).
- `apps-script/Code.gs`: el código del backend para pegar en Apps Script.

## Paso 1 · Hoja y Apps Script (10 min)
1. En Google Drive crea una hoja nueva: **"Propuestas Chocó 2026 · Ronda 2"**.
2. En la hoja: **Extensiones → Apps Script**. Borra lo que haya y pega todo el contenido de `apps-script/Code.gs`. Guarda.
3. Arriba, en el selector de funciones, elige **setup** y dale **Ejecutar**. Acepta los permisos. Esto crea las pestañas `fichas` e `historial`.
4. **Implementar → Nueva implementación**:
   - Tipo: **Aplicación web**
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
5. Copia la **URL de la aplicación web** (termina en `/exec`).

## Paso 2 · Conectar la página
1. Abre `index.html` y busca, casi al inicio del script:
   ```js
   const API = '';
   const ADMIN_PIN = '2611';
   ```
2. Pega la URL entre las comillas de `API` y cambia el PIN de coordinación si quieres.
3. Mientras `API` esté vacío, la página funciona en **modo prueba** (guarda solo en el celular).

## Paso 3 · GitHub Pages
1. En GitHub (cuenta `mateolaverde97-byte`): **New repository** → nombre `mision-choco-r2` → **Public** → Create.
2. **Add file → Upload files** → sube `index.html` (y si quieres `README.md` y la carpeta `apps-script`). Commit.
3. **Settings → Pages** → Source: **Deploy from a branch** → Branch: **main** / **(root)** → Save.
4. En 1–2 minutos queda en: `https://mateolaverde97-byte.github.io/mision-choco-r2/`

## Si cambias el código de Apps Script
**Implementar → Gestionar implementaciones → editar (lápiz) → Versión: Nueva versión → Implementar.** La URL se mantiene.

## Cómo funciona
- Cada ficha es una fila en la pestaña `fichas`; cada guardado queda en `historial`.
- Las fichas se guardan solas mientras escriben. Si dos personas editan a la vez, la página avisa y deja elegir qué versión conservar.
- "Enviar a Matt L." marca la ficha como enviada en la hoja y abre WhatsApp con el resumen y un código (sirve de respaldo en "Traer propuestas").
- Coordinación (con PIN): mapa de propuestas, lado a lado por bloque, lista de compras consolidada (descarga CSV) y traer propuestas.
