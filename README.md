# Imprime Cuadre

Herramienta web para subir el Excel del cuadre diario y descargar los PDFs por día y los totales generales.

**🌐 En vivo:** https://juanfnmb1.github.io/imprime-cuadre/

## Cómo se usa

1. Abrir la página.
2. Subir el Excel (con las hojas `Totales` y `Transacciones Organizadas`).
3. Revisar la vista previa.
4. Descargar el PDF de cada día, el de Totales Generales, o todos a la vez (ZIP).

Los montos se toman tal cual están en el Excel — no se recalculan.

## Desarrollo

```bash
npm install
npm run dev      # servidor de desarrollo local
npm run build    # genera /dist para producción
npm run preview  # sirve el build local
```

## Despliegue (GitHub Pages)

El workflow `.github/workflows/deploy.yml` despliega automáticamente al hacer push a `main`.

Para activar GitHub Pages la primera vez:
1. En el repo: **Settings → Pages → Build and deployment → Source = GitHub Actions**.
2. Hacer push a `main`.
3. La URL aparecerá en la pestaña Actions y en Settings → Pages.
