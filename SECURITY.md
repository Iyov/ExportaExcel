# Política de Seguridad

## Versiones Soportadas

| Versión | Soporte          |
| ------- | ---------------- |
| 2.0.x   | ✅ Soportada     |
| < 2.0   | ❌ No soportada  |

## Reportar Vulnerabilidades

Si descubres una vulnerabilidad de seguridad, por favor **NO abras un issue público**.

### Cómo Reportar

1. **Email**: Envía un correo a [security@tu-dominio.com] con:
   - Descripción detallada de la vulnerabilidad
   - Pasos para reproducir el problema
   - Impacto potencial
   - Sugerencias de solución (si las tienes)

2. **Respuesta**: Recibirás confirmación en 48 horas

3. **Resolución**: Trabajaremos en un fix y te notificaremos cuando esté disponible

## Mejores Prácticas de Seguridad

### Para Usuarios

#### 1. Protege tus Credenciales
```bash
# ✅ Correcto: database.ini en .gitignore
database.ini

# ❌ NUNCA hagas esto:
git add database.ini
git commit -m "agrega credenciales"
```

#### 2. Usa Conexiones Seguras
- Configura SSL/TLS para conexiones a SQL Server
- Usa credenciales con permisos mínimos necesarios
- Rota contraseñas regularmente

#### 3. Mantén el Código Actualizado
```bash
# Actualizar dependencias
pip install --upgrade -r requirements.txt

# Verificar vulnerabilidades conocidas
pip-audit
```

### Para Desarrolladores

#### 1. Prevención de SQL Injection

```python
# ✅ Correcto: Usa parámetros
query = "SELECT * FROM Tabla WHERE id = ?"
cursor.execute(query, (id_valor,))

# ❌ NUNCA hagas esto:
query = f"SELECT * FROM Tabla WHERE id = {id_valor}"
cursor.execute(query)
```

#### 2. No Loguees Información Sensible

```python
# ✅ Correcto
logger.info("Conexión exitosa a la base de datos")

# ❌ Evitar
logger.info(f"Conectado con usuario {username} y password {password}")
```

#### 3. Valida Todas las Entradas

```python
# ✅ Correcto
if not validate_date_range(desde, hasta):
    raise ValueError("Rango de fechas inválido")
```

## Vulnerabilidades Conocidas

### Versión 1.x (No Soportada)

- **SQL Injection**: Queries concatenadas con f-strings
- **Credenciales expuestas**: database.ini en repositorio Git
- **Conexiones no cerradas**: Fugas de recursos

**Solución**: Actualizar a versión 2.0+

### Versión 2.0+

No hay vulnerabilidades conocidas. Si encuentras alguna, repórtala siguiendo el proceso anterior.

## Recursos

- [OWASP Top 10](https://owasp.org/www-project-top-ten/)
- [Python Security](https://python.readthedocs.io/en/stable/library/security_warnings.html)
- [SQL Injection Prevention](https://cheatsheetseries.owasp.org/cheatsheets/SQL_Injection_Prevention_Cheat_Sheet.html)

---

**Última actualización**: Febrero 2026
