# Análisis Cruzado P4 - Servicio Electrónico

## Descripción General

Este documento explica la metodología utilizada para generar el análisis cruzado de la pregunta **P4 - Servicio Electrónico** con todas las variables demográficas y de clasificación.

## Problema a Resolver

La pregunta P4 permite que los usuarios seleccionen un único servicio electrónico de las siguientes opciones:
- `a. RTU`
- `b. FEL`
- `c. Aduanas sin papeles`
- `d. Agencia Virtual`
- `e. Otros`

A diferencia de P3, **P4 no permite selecciones múltiples**, por lo que cada registro tiene un único valor de P4.

## Metodología de Normalización

### 1. Sin Normalización Necesaria

Como P4 no tiene combinaciones múltiples, no se requiere normalización. Los valores se utilizan directamente de la columna `P4 - Servicio Electrónico`.

```python
# P4 no requiere normalización (no tiene combinaciones múltiples)
df['P4 - Servicio Electrónico']  # Se usa directamente
```

### Ejemplo de Valores

| Valor Original | Uso en Análisis |
|---------------|-----------------|
| `a. RTU` | `a. RTU` |
| `b. FEL` | `b. FEL` |
| `c. Aduanas sin papeles` | `c. Aduanas sin papeles` |
| `d. Agencia Virtual` | `d. Agencia Virtual` |
| `e. Otros` | `e. Otros` |

## Metodología de Cálculo

### Principio Fundamental

**Cada opción se cuenta directamente mediante comparación exacta (==), ya que no hay combinaciones múltiples.**

### Proceso de Cálculo

#### 1. Conteo por Opción Individual

Para cada opción principal, se cuenta directamente:

```python
rtu_total = len(df[df['P4 - Servicio Electrónico'] == 'a. RTU'])
fel_total = len(df[df['P4 - Servicio Electrónico'] == 'b. FEL'])
aduanas_total = len(df[df['P4 - Servicio Electrónico'] == 'c. Aduanas sin papeles'])
agencia_total = len(df[df['P4 - Servicio Electrónico'] == 'd. Agencia Virtual'])
otros_total = len(df[df['P4 - Servicio Electrónico'] == 'e. Otros'])
```

#### 2. Distribución de Registros

**a. RTU**: 183 registros
**b. FEL**: 26 registros
**c. Aduanas sin papeles**: 24 registros
**d. Agencia Virtual**: 280 registros
**e. Otros**: 15 registros

**Total**: 183 + 26 + 24 + 280 + 15 = 528 registros

### 3. Cruces con Otras Variables

Para cada cruce con otra variable (ej: Género, Rango de Edad, etc.), se aplica comparación directa:

```python
# Ejemplo: Cruce P4 vs Género
rtu_h = len(df[
    (df['P4 - Servicio Electrónico'] == 'a. RTU') & 
    (df['P37 - Género'] == 'H')
])
```

**Importante**: Cada registro se cuenta una sola vez en su opción correspondiente de P4.

## Estructura del Archivo Excel Generado

### Filas del Análisis

1. **Fila 1**: Título "P4 - Servicio Electrónico" (alineado a la izquierda)
2. **Fila 2**: Fila vacía con fondo gris (formato)
3. **Fila 3**: Encabezados principales de variables
4. **Fila 4**: Sub-encabezados (categorías de cada variable)
5. **Filas 5-9**: Datos de las 5 opciones de P4
   - Fila 5: `a. RTU`
   - Fila 6: `b. FEL`
   - Fila 7: `c. Aduanas sin papeles`
   - Fila 8: `d. Agencia Virtual`
   - Fila 9: `e. Otros`
6. **Fila 10**: Totales
7. **Filas 11-12**: Filas vacías (separación)
8. **Filas 13-17**: Tabla de porcentajes (misma estructura)

### Columnas del Análisis

1. **Columna 1**: Nombre de la opción de P4
2. **Columna 2**: TOTAL general
3. **Columnas 3-5**: P37 Género (H, M, No deseo responder)
4. **Columnas 6-10**: Rango de edad (18-25, 26-35, 36-45, 46-60, Más de 61)
5. **Columnas 11-21**: P40 Nivel académico (11 categorías)
6. **Columnas 22-38**: P39 Idiomas (17 categorías)
7. **Columnas 39-60**: P44 Oficina/Agencia/Delegación (22 departamentos)
8. **Columnas 61-71**: P44.1 Aduana (11 aduanas)
9. **Columnas 72-86**: P9 Personería (15 categorías)
10. **Columnas 87-91**: P38 Etnia (5 categorías)
11. **Columnas 92-95**: Oficina/Agencia/Delegación por Región (4 regiones)
12. **Columnas 96-99**: Aduana por Región (4 regiones)

## Validación de Resultados

### Verificación de Totales

- **Total de registros con P4**: 528
- **a. RTU**: 183 registros (34.7%)
- **b. FEL**: 26 registros (4.9%)
- **c. Aduanas sin papeles**: 24 registros (4.5%)
- **d. Agencia Virtual**: 280 registros (53.0%)
- **e. Otros**: 15 registros (2.8%)

**Nota**: La suma de las cinco opciones (183 + 26 + 24 + 280 + 15 = 528) coincide exactamente con el total de registros porque cada registro tiene un único valor de P4.

### Verificación de Consistencia

| Opción | Total | Verificación |
|--------|-------|--------------|
| `a. RTU` | 183 | ✓ |
| `b. FEL` | 26 | ✓ |
| `c. Aduanas sin papeles` | 24 | ✓ |
| `d. Agencia Virtual` | 280 | ✓ |
| `e. Otros` | 15 | ✓ |
| **TOTAL** | **528** | ✓ |

## Ejemplo de Cálculo Detallado

### Caso: Registro Individual

**Registro**: ID 7
- **P4**: `a. RTU`
- **Género**: `H`
- **Rango de Edad**: `26 - 35`

**Conteo en el análisis**:
- Se cuenta en `a. RTU + TOTAL`: ✓
- Se cuenta en `a. RTU + H`: ✓
- Se cuenta en `a. RTU + 26 - 35`: ✓

### Caso: Registro con Agencia Virtual

**Registro**: Con `d. Agencia Virtual`
- **P4**: `d. Agencia Virtual`
- **Género**: `M`

**Conteo en el análisis**:
- Se cuenta en `d. Agencia Virtual`: ✓
- Se cuenta en `d. Agencia Virtual + M`: ✓
- No se cuenta en otras opciones de P4 (solo una opción por registro)

## Tabla de Porcentajes

La segunda tabla en el archivo Excel muestra los porcentajes de cada opción de P4:

### Cálculo de Porcentajes

1. **Porcentaje sobre el total general**: Cada opción se calcula como porcentaje del total de registros con P4 (528).

2. **Porcentaje por fila**: Para cada cruce con otras variables, el porcentaje se calcula sobre el total de esa fila (opción de P4).

3. **Redondeo**: Los porcentajes se redondean a enteros, con la regla de que valores >= 0.5 se redondean hacia arriba.

### Ejemplo de Porcentajes

- **a. RTU**: 183 / 528 = 34.66% → 35%
- **d. Agencia Virtual**: 280 / 528 = 53.03% → 53%

## Ventajas de esta Metodología

1. **Simplicidad**: Al no haber combinaciones múltiples, el cálculo es directo y simple.

2. **Precisión**: Cada registro se cuenta exactamente una vez en su opción correspondiente.

3. **Consistencia**: Los totales son exactos y verificables sin ambigüedad.

4. **Claridad**: La estructura es más clara al no tener que manejar combinaciones.

## Uso del Script

```bash
# Uso básico (archivos por defecto)
python3 P4-Cruzado.py

# Especificar archivo de entrada
python3 P4-Cruzado.py V3.xlsx

# Especificar entrada y salida
python3 P4-Cruzado.py V3.xlsx P4-Cruzado.xlsx
```

## Notas Técnicas

- El script utiliza `pandas` para el procesamiento de datos
- `openpyxl` para la generación del archivo Excel con formato
- La comparación directa (`==`) se utiliza en lugar de `str.contains()` ya que P4 no tiene combinaciones
- Los bordes y formatos siguen el mismo estilo que P3-Cruzado.xlsx
- La tabla de porcentajes incluye el signo "%" y valores redondeados a enteros

## Diferencias con P3

| Aspecto | P3 | P4 |
|---------|----|----|
| **Combinaciones múltiples** | Sí | No |
| **Normalización** | Requerida | No requerida |
| **Método de conteo** | `str.contains()` | `==` (comparación directa) |
| **Número de opciones** | 3 | 5 |
| **Suma de opciones** | > Total (por combinaciones) | = Total (sin combinaciones) |

## Archivos Relacionados

- **Script Python**: `P4-Cruzado.py`
- **Archivo de entrada**: `V3.xlsx`
- **Archivo de salida**: `P4-Cruzado.xlsx`
- **Documentación**: `P4-Cruzado.md` (este archivo)

