# Changelog

Todas las mejoras y cambios importantes de este proyecto se documentan en este archivo.

El formato sigue las recomendaciones de:
https://keepachangelog.com/es-ES/1.0.0/

---

## [1.0.0] - 2026-05-28

### 🎉 Primera versión estable

#### ✅ Añadido
- Conversión de HTML a Excel (.xlsx)
- Soporte para:
  - Párrafos y texto formateado
  - Encabezados (h1–h6)
  - Listas anidadas (ul, ol)
  - Tablas complejas (rowspan / colspan)
  - Enlaces internos y externos
  - Imágenes (base64, URL, rutas locales)
- Sistema de logging configurable (debug / quiet)
- Detección de archivos en uso
- Validación de entrada y salida

---

### 🖼️ Imágenes (motor avanzado)
- Escalado automático de imágenes
- Soporte de:
  - `--image-max-width`
  - `--image-max-height`
- Modos de ajuste:
  - `contain`
  - `cover`
  - `stretch`
- `--image-padding` (espaciado vertical)
- `--image-background` (color de fondo)
- `--image-align` (left, center, right)
- `--image-border` (borde exterior sin rejilla)

---

### 📊 Excel / Layout
- Fusión dinámica de columnas
- Altura de filas autoajustada
- Formato especial para:
  - Código (monoespaciado)
  - Encabezados
  - Tablas
- Soporte para estilos básicos:
  - negrita
  - cursiva
  - subrayado

---

### ⚙️ CLI (línea de comandos)
- Parser completo con argparse
- Opciones:
  - `--columns`
  - `--column-width`
  - `--no-images`
  - `--no-tables`
  - `--no-links`
- `--help-full` con documentación extendida

---

### 🔐 Licencia
- GPL-3.0-or-later

---

## 🔮 Próximas mejoras (ideas)

- Caché de imágenes para evitar duplicados
- Soporte para estilos CSS básicos
- Alineación vertical de imágenes
- Soporte para bordes personalizables
- Exportación a múltiples hojas Excel