# Análisis Cruzado P3 - Medios SAT Utilizados

## Descripción General

Este documento explica la metodología utilizada para generar el análisis cruzado de la pregunta **P3 - Medios SAT Utilizados** con todas las variables demográficas y de clasificación.

## Problema a Resolver

La pregunta P3 permite que los usuarios seleccionen múltiples opciones simultáneamente, generando combinaciones como:
- Opciones individuales: `a. Presencial`, `b. Contact Center`, `c. Servicios Electrónicos`
- Combinaciones de 2 opciones: `a. Presencial, b. Contact Center`
- Combinaciones de 3 opciones: `a. Presencial, b. Contact Center, c. Servicios Electrónicos`

Además, las combinaciones pueden aparecer en diferentes órdenes en los datos originales (ej: `b. Contact Center, a. Presencial` vs `a. Presencial, b. Contact Center`).

## Metodología de Normalización

### 1. Función de Normalización

Se creó una función `normalizar_p3()` que:

1. **Detecta las opciones presentes** en cada respuesta:
   ```python
   tiene_presencial = 'a. Presencial' in valor_str
   tiene_contact = 'b. Contact Center' in valor_str
   tiene_electronicos = 'c. Servicios Electrónicos' in valor_str
   ```

2. **Construye la combinación normalizada** en un orden estándar:
   ```python
   opciones = []
   if tiene_presencial:
       opciones.append('a. Presencial')
   if tiene_contact:
       opciones.append('b. Contact Center')
   if tiene_electronicos:
       opciones.append('c. Servicios Electrónicos')
   
   return ', '.join(opciones)
   ```

3. **Resultado**: Todas las variaciones del mismo conjunto de opciones se agrupan en una sola combinación normalizada.

### Ejemplo de Normalización

| Valor Original | Valor Normalizado |
|---------------|-------------------|
| `a. Presencial` | `a. Presencial` |
| `b. Contact Center, a. Presencial` | `a. Presencial, b. Contact Center` |
| `a. Presencial, b. Contact Center` | `a. Presencial, b. Contact Center` |
| `c. Servicios Electrónicos, a. Presencial, b. Contact Center` | `a. Presencial, b. Contact Center, c. Servicios Electrónicos` |

## Metodología de Cálculo

### Principio Fundamental

**Cada opción individual debe incluir TODOS los registros que la contengan, independientemente de si está sola o en combinación con otras opciones.**

### Proceso de Cálculo

#### 1. Conteo por Opción Individual

Para cada opción principal (a. Presencial, b. Contact Center, c. Servicios Electrónicos), se cuenta:

```python
presencial_total = len(df[df['P3_norm'].str.contains('a. Presencial', na=False)])
contact_total = len(df[df['P3_norm'].str.contains('b. Contact Center', na=False)])
electronicos_total = len(df[df['P3_norm'].str.contains('c. Servicios Electrónicos', na=False)])
```

#### 2. Desglose de Inclusión

Cada opción incluye:

**a. Presencial (699 registros totales):**
- 633 registros con solo `a. Presencial`
- 41 registros con `a. Presencial, b. Contact Center`
- 22 registros con `a. Presencial, c. Servicios Electrónicos`
- 3 registros con `a. Presencial, b. Contact Center, c. Servicios Electrónicos`
- **Total: 633 + 41 + 22 + 3 = 699**

**b. Contact Center (176 registros totales):**
- 128 registros con solo `b. Contact Center`
- 41 registros con `a. Presencial, b. Contact Center`
- 4 registros con `b. Contact Center, c. Servicios Electrónicos`
- 3 registros con `a. Presencial, b. Contact Center, c. Servicios Electrónicos`
- **Total: 128 + 41 + 4 + 3 = 176**

**c. Servicios Electrónicos (528 registros totales):**
- 499 registros con solo `c. Servicios Electrónicos`
- 22 registros con `a. Presencial, c. Servicios Electrónicos`
- 4 registros con `b. Contact Center, c. Servicios Electrónicos`
- 3 registros con `a. Presencial, b. Contact Center, c. Servicios Electrónicos`
- **Total: 499 + 22 + 4 + 3 = 528**

### 3. Cruces con Otras Variables

Para cada cruce con otra variable (ej: Género, Rango de Edad, etc.), se aplica el mismo principio:

```python
# Ejemplo: Cruce P3 vs Género
presencial_h = len(df[
    (df['P3_norm'].str.contains('a. Presencial', na=False)) & 
    (df['P37 - Género'] == 'H')
])
```

**Importante**: Un registro con combinación `a. Presencial, b. Contact Center` y Género='H' se cuenta en:
- `a. Presencial + H`
- `b. Contact Center + H`

## Estructura del Archivo Excel Generado

### Filas del Análisis

1. **Fila 1**: Fondo gris (formato)
2. **Fila 2**: Encabezados principales de variables
3. **Fila 3**: Sub-encabezados (categorías de cada variable)
4. **Filas 4-6**: Datos de las 3 opciones de P3
   - Fila 4: `a. Presencial`
   - Fila 5: `b. Contact Center`
   - Fila 6: `c. Servicios Electrónicos`
5. **Fila 7**: Totales

### Columnas del Análisis

1. **Columna 1**: Nombre de la opción de P3
2. **Columna 2**: TOTAL general
3. **Columnas 3-5**: P37 Género (H, M, No deseo responder)
4. **Columnas 6-10**: Rango de edad (18-25, 26-35, 36-45, 46-60, Más de 61)
5. **Columnas 11-21**: P40 Nivel académico (11 categorías)
6. **Columnas 22-38**: P39 Idiomas (17 categorías)
7. **Columnas 39-60**: P44 Oficina/Agencia/Delegación (22 departamentos)
8. **Columnas 61-71**: P44.1 Aduana (11 aduanas)
9. **Columnas 72-86**: P9 Personería (15 categorías)
10. **Columnas 87-91**: P38 Etnia (5 categorías)
11. **Columnas 92-94**: P3 Medios SAT utilizados (distribución)

## Validación de Resultados

### Verificación de Totales

- **Total de registros**: 1,330
- **a. Presencial**: 699 registros (52.6%)
- **b. Contact Center**: 176 registros (13.2%)
- **c. Servicios Electrónicos**: 528 registros (39.7%)

**Nota**: La suma de las tres opciones (699 + 176 + 528 = 1,403) es mayor que el total de registros (1,330) porque algunos registros tienen múltiples opciones y se cuentan en cada una.

### Verificación de Combinaciones

| Combinación | Cantidad | Incluida en |
|------------|----------|-------------|
| Solo `a. Presencial` | 633 | a. Presencial |
| Solo `b. Contact Center` | 128 | b. Contact Center |
| Solo `c. Servicios Electrónicos` | 499 | c. Servicios Electrónicos |
| `a. Presencial, b. Contact Center` | 41 | a. Presencial + b. Contact Center |
| `a. Presencial, c. Servicios Electrónicos` | 22 | a. Presencial + c. Servicios Electrónicos |
| `b. Contact Center, c. Servicios Electrónicos` | 4 | b. Contact Center + c. Servicios Electrónicos |
| `a. Presencial, b. Contact Center, c. Servicios Electrónicos` | 3 | Las tres opciones |

## Ejemplo de Cálculo Detallado

### Caso: Registro con Combinación

**Registro**: ID 7
- **P3 Original**: `a. Presencial, b. Contact Center`
- **P3 Normalizado**: `a. Presencial, b. Contact Center`
- **Género**: `H`

**Conteo en el análisis**:
- Se cuenta en `a. Presencial + TOTAL`: ✓
- Se cuenta en `a. Presencial + H`: ✓
- Se cuenta en `b. Contact Center + TOTAL`: ✓
- Se cuenta en `b. Contact Center + H`: ✓

### Caso: Registro con Tres Opciones

**Registro**: Con las tres opciones seleccionadas
- **P3 Normalizado**: `a. Presencial, b. Contact Center, c. Servicios Electrónicos`

**Conteo en el análisis**:
- Se cuenta en `a. Presencial`: ✓
- Se cuenta en `b. Contact Center`: ✓
- Se cuenta en `c. Servicios Electrónicos`: ✓
- Se cuenta en todos los cruces correspondientes de cada opción

## Ventajas de esta Metodología

1. **Inclusión completa**: Todos los registros que seleccionaron una opción se incluyen, incluso si también seleccionaron otras.

2. **Consistencia**: Los totales son consistentes y verificables.

3. **Flexibilidad**: Permite analizar cada opción por separado sin perder información de las combinaciones.

4. **Normalización**: Agrupa automáticamente variaciones del mismo conjunto de opciones.

## Uso del Script

```bash
# Uso básico (archivos por defecto)
python3 P3-Cruzado.py

# Especificar archivo de entrada
python3 P3-Cruzado.py V3.xlsx

# Especificar entrada y salida
python3 P3-Cruzado.py V3.xlsx Analisis_Cruzado_P3.xlsx
```

## Notas Técnicas

- El script utiliza `pandas` para el procesamiento de datos
- `openpyxl` para la generación del archivo Excel con formato
- La función `str.contains()` se utiliza para detectar si una combinación incluye una opción específica
- Los bordes y formatos siguen el estilo del archivo de ejemplo proporcionado

## Archivos Relacionados

- **Script Python**: `P3-Cruzado.py`
- **Archivo de entrada**: `V3.xlsx`
- **Archivo de salida**: `Analisis_Cruzado_P3.xlsx`
- **Documentación**: `P3-Cruzado.md` (este archivo)

