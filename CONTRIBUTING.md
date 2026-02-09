# Guía de Contribución

¡Gracias por tu interés en contribuir a ExportaExcel! 🎉

## Cómo Contribuir

### 1. Reportar Bugs

Si encuentras un error:
- Verifica que no exista un issue similar
- Abre un nuevo issue describiendo:
  - Qué esperabas que pasara
  - Qué pasó realmente
  - Pasos para reproducir el error
  - Versión de Python y dependencias

### 2. Sugerir Mejoras

Para proponer nuevas funcionalidades:
- Abre un issue con la etiqueta `enhancement`
- Describe claramente la mejora propuesta
- Explica por qué sería útil

### 3. Enviar Pull Requests

1. **Fork** el repositorio
2. **Crea una rama** para tu cambio:
   ```bash
   git checkout -b feature/mi-mejora
   ```
3. **Haz tus cambios** siguiendo las convenciones del código
4. **Prueba** que todo funcione correctamente
5. **Commit** con un mensaje descriptivo:
   ```bash
   git commit -m "Agrega validación de fechas en BD.py"
   ```
6. **Push** a tu fork:
   ```bash
   git push origin feature/mi-mejora
   ```
7. **Abre un Pull Request** describiendo los cambios

## Estándares de Código

### Estilo Python
- Usa nombres descriptivos en español para variables y funciones
- Sigue PEP 8 (puedes usar `black` para formateo automático)
- Agrega comentarios para lógica compleja
- Mantén funciones cortas y enfocadas

### Ejemplo
```python
# ✅ Bien
def obtener_datos_contrato(id_contrato):
    """Obtiene los datos de un contrato desde la BD."""
    query = "SELECT * FROM Contrato WHERE IdContrato = ?"
    return ejecutar_query(query, (id_contrato,))

# ❌ Evitar
def gdc(x):
    return ejecutar_query(f"SELECT * FROM Contrato WHERE IdContrato = {x}")
```

## Convenciones de Commit

Usa mensajes claros y descriptivos:

```bash
# Buenos ejemplos
git commit -m "Agrega validación de fechas en validators.py"
git commit -m "Corrige error de conexión en BD.py"
git commit -m "Actualiza README con instrucciones de instalación"

# Evitar
git commit -m "fix"
git commit -m "cambios"
git commit -m "update"
```

## Proceso de Review

Tu Pull Request será revisado por un maintainer. Espera:
- Feedback constructivo
- Posibles solicitudes de cambios
- Aprobación o explicación si no se acepta

## Configuración de Desarrollo

```bash
# Clonar tu fork
git clone https://github.com/FcoBarrientos/ExportaExcel.git
cd ExportaExcel

# Instalar dependencias
pip install -r requirements.txt

# Configurar base de datos de prueba
cp database.ini.example database.ini
# Editar database.ini con credenciales de desarrollo

# Verificar que todo funcione
python check_setup.py
```

## Preguntas

¿Tienes dudas? Abre un issue con la etiqueta `question`.

---

**Gracias por contribuir a ExportaExcel!** 🚀
