# html2xlsx

Convierte documentos HTML en archivos Excel (.xlsx) manteniendo su estructura original.

## ✨ Características

- ✅ Encabezados y párrafos
- ✅ Listas anidadas (ul / ol)
- ✅ Tablas complejas (rowspan / colspan)
- ✅ Imágenes (base64, URL, locales)
- ✅ Enlaces internos y externos
- ✅ Control avanzado de layout de imágenes

---

## 🚀 Instalación

1. Clona o descarga este repositorio
2. Instala dependencias:

```bash
pip install -r requirements.txt
```

---

## ▶️ Uso básico

```bash
python html2xlsx.py archivo.html
```

Esto generará automáticamente:

```
archivo.xlsx
```

---

## ⚙️ Opciones

### Generales

| Opción | Descripción |
|------|-------------|
| `--columns N` | Número de columnas |
| `--column-width N` | Ancho de columnas |
| `--no-images` | Desactiva imágenes |
| `--no-tables` | Desactiva tablas |
| `--no-links` | Desactiva enlaces |

---

### Imágenes

| Opción | Descripción |
|------|-------------|
| `--image-max-width N` | Ancho máximo |
| `--image-max-height N` | Alto máximo |
| `--image-fit` | contain, cover, stretch |
| `--image-padding N` | Espacio vertical |
| `--image-background HEX` | Color fondo |
| `--image-align` | left, center, right |
| `--image-border` | Añade borde |

---

## 📌 Ejemplos

### Conversión básica

```bash
python html2xlsx.py archivo.html
```

---

### Layout profesional

```bash
python html2xlsx.py archivo.html \
    --image-align center \
    --image-background F8F8F8 \
    --image-border
```

---

### Control de tamaño de imágenes

```bash
python html2xlsx.py archivo.html \
    --image-max-width 800 \
    --image-max-height 400
```

---

### Sin imágenes ni tablas

```bash
python html2xlsx.py archivo.html \
    --no-images \
    --no-tables
```

---

## 📄 Licencia

Este proyecto está licenciado bajo GNU GPL v3.

---

## 👨‍💻 Autor

Antonio Teodomiro Márquez Muñoz (Naidel)

📧 atmarquez@gmail.com  
💰 https://paypal.me/atmarquez  
