# Proyecto Parcial 2 – Ciudadanía Digital I: Cultura Digital

Material listo para cada paso del proyecto. Lo que tienes que hacer tú mismo (crear cuentas, encuestar a personas reales y publicar) está marcado con **TÚ**.

---

## Paso 1 – La nube y el formulario (5 pts)

### 1.1 Carpeta en la nube · **TÚ**
1. En Google Drive: **Nuevo → Carpeta** → `Proyecto_Competencias_Digitales`
2. Dentro, crea esta estructura (ordenada y jerárquica):
   ```
   Proyecto_Competencias_Digitales/
   ├── 01_Formulario/
   ├── 02_Base_de_Datos/        ← aquí va Apellido_Nombre_Proyecto.xlsx
   ├── 03_Graficos/             ← imágenes PNG de los gráficos
   └── 04_Diario_Online/        ← imágenes de fondo e íconos
   ```
3. Clic derecho en la carpeta → **Compartir → Cualquier persona con el enlace → Lector** (la rúbrica pide carpeta pública).
4. Toma una **captura de pantalla** de la carpeta para tu planeación o tus apuntes.

### 1.2 Formulario (Google Forms) · cópialo tal cual
**Título:** Hábitos de estudio y uso de redes sociales antes de dormir

**Descripción:**
> Encuesta escolar para el proyecto de Ciudadanía Digital del CEUM. Es **anónima**: no se piden nombres, correos ni teléfonos. Tus respuestas solo se usarán con fines educativos y estadísticos. ¡Gracias por participar!

| # | Pregunta | Tipo | Configuración |
|---|----------|------|---------------|
| 1 | Edad | Respuesta corta | Obligatoria · ⋮ → **Validación de respuestas → Número → Número entero** |
| 2 | ¿Estudias más de 2 horas al día? | Opción múltiple | Opciones: `Sí` / `No` · Obligatoria |
| 3 | ¿Usas TikTok o Instagram antes de dormir? | Opción múltiple | Opciones: `Sí` / `No` · Obligatoria |

⚠️ **Ética (Regla de Oro):** en ⚙️ Configuración → Respuestas, deja **desactivado** "Recopilar direcciones de correo electrónico" y "Limitar a 1 respuesta" (esta última obliga a iniciar sesión).

**TÚ:** comparte el enlace y consigue **50 respuestas reales**.

---

## Paso 2 – Base de datos y limpieza (10 pts)

El archivo `Apellido_Nombre_Proyecto.xlsx` ya tiene la estructura profesional:

- **Columnas = campos:** Folio · Edad · ¿Estudias más de 2 horas al día? · ¿Usas TikTok o Instagram antes de dormir?
- **Filas = registros:** filas 2 a 51 (exactamente 50) con un folio anónimo `R01`–`R50`.
- **Validación de datos:** la edad solo acepta números enteros (10–99) y las preguntas solo aceptan `Sí` o `No`, así que no se pueden colar respuestas corruptas.
- **Las celdas vacías se marcan en rojo** automáticamente.

**Qué haces tú:**
1. **Renombra** el archivo con tu apellido y tu nombre, por ejemplo `Pérez_Ana_Proyecto.xlsx`.
2. En Forms → **Respuestas → ⋮ → Descargar respuestas (.csv)** y ábrelo en Excel.
3. En ese archivo descargado: selecciona toda la tabla → **Datos → Quitar duplicados**, **dejando marcadas todas las columnas, incluida "Marca temporal"**. Así se eliminan los envíos repetidos de verdad. *(Con solo 3 preguntas es normal que dos personas distintas respondan igual, por ejemplo "16, Sí, Sí". Esas filas NO son duplicados: son personas diferentes.)*
4. Revisa a mano que no haya vacíos ni edades raras ("dieciséis", "16 años"). Corrígelas o elimina esa fila.
5. Copia **solo** las columnas Edad y las dos de Sí/No (no copies la "Marca temporal") y pégalas en **B2** del archivo del proyecto con **Pegado especial → Valores**, para que no se pierda el formato.
6. Si quedan menos de 50 respuestas válidas, consigue más hasta completar 50 (y si sobran, quédate con las primeras 50).
7. Revisa la **Tabla 3**: debe decir **"✔ Base completa: 50 registros válidos"**.

---

## Paso 3 – Fórmulas y automatización (15 pts)

Todo está a la derecha de los datos (columnas F–H) y se calcula solo. No hay ningún número escrito a mano.

| Qué calcula | Celda | Fórmula (Excel en español) |
|---|---|---|
| Sí estudian +2 h | G3 | `=CONTAR.SI(C2:C51,F3)` |
| No estudian +2 h | G4 | `=CONTAR.SI(C2:C51,F4)` |
| Sí usan redes | G9 | `=CONTAR.SI(D2:D51,F9)` |
| No usan redes | G10 | `=CONTAR.SI(D2:D51,F10)` |
| Totales | G5, G11 | `=SUMA(G3:G4)` |
| Porcentajes | H3… | `=SI($G$5=0,0,G3/$G$5)` |
| Edad promedio | H14 | `=PROMEDIO(B2:B51)` |
| Edad máxima | H15 | `=MAX(B2:B51)` |
| Edad mínima | H16 | `=MIN(B2:B51)` |
| Total de encuestados | H17 | `=CONTARA(B2:B51)` |
| Control de vacíos | H18 | `=CONTAR.BLANCO(B2:D51)` |
| Control de errores en edad | H19 | `=CONTARA(B2:B51)-CONTAR(B2:B51)` |
| Cruce de variables (extra) | H24… | `=CONTAR.SI.CONJUNTO(D2:D51,F24,C2:C51,G24)` |

✅ El criterio de `CONTAR.SI` apunta a una **celda** (`F3`, que contiene "Sí"), no a un valor escrito a mano, tal como pide la rúbrica.
ℹ️ El archivo se generó con los nombres de función en inglés; tu Excel en español los muestra automáticamente como `CONTAR.SI`, `PROMEDIO`, etc.
ℹ️ Mientras la tabla esté vacía, `PROMEDIO`, `MAX` y `MIN` pueden mostrar `#¡DIV/0!` o `0`. Es normal: se corrige solo al pegar los datos.

### Gráficos circulares (ya insertados)
- **Gráfico 1:** Porcentaje de alumnos que usan TikTok o Instagram antes de dormir
- **Gráfico 2:** Porcentaje de alumnos que estudian más de 2 horas al día

Se actualizan solos. **Para exportarlos en alta definición:**
1. Clic en el gráfico → agrándalo (arrastra la esquina) para que tenga más resolución.
2. Clic derecho → **Guardar como imagen… → PNG** → guárdalo en `03_Graficos/`.
   *(Si tu Excel no tiene esa opción: copia el gráfico y pégalo en PowerPoint, y ahí usa clic derecho → Guardar como imagen.)*

---

## Paso 4 – Diario Online en Google Sites (20 pts)

### 4.1 Estructura del sitio (menú con íconos)
Crea un sitio en sites.google.com con estas **páginas** (aparecerán como menú):

| Página | Ícono sugerido | Contenido |
|---|---|---|
| 🏠 Inicio | casa | Portada, presentación y objetivo |
| 📋 Metodología | portapapeles | Cómo se hizo la encuesta y la ética |
| 📊 Resultados | gráfica | Gráficos 1 y 2 con su análisis |
| 🌍 Contraste | globo | Comparación con fuentes oficiales |
| ⚖️ Conclusión ética | balanza | Conclusión y recomendaciones |

- **Imágenes de fondo libres de derechos:** en pixabay.com busca "teenager phone night", "student studying", "smartphone bed". Úsalas como imagen de encabezado (**Tema → Encabezado → Cambiar imagen**).
- **Íconos:** Google Sites no tiene una biblioteca de íconos. Pon el emoji al inicio del nombre de cada página o descarga íconos gratuitos (flaticon.com o el buscador de Canva) y súbelos como imagen.
- **Gráficos:** **Insertar → Imágenes → Subir** el PNG. No lo estires desde las esquinas laterales, para que no se deforme.

### 4.2 Textos listos para pegar
> Reemplaza todo lo que esté entre **[corchetes]** con los números de **tu** Excel (Tablas 1 a 4).

---

#### 🏠 INICIO
**Título:** Pantallas, sueño y estudio: ¿qué hacen los jóvenes antes de dormir?

Bienvenidos a este Diario Online, creado para la materia Ciudadanía Digital I: Cultura Digital del Colegio de Estudios del Mayab. Aquí presentamos los resultados de una encuesta anónima aplicada a 50 personas sobre dos hábitos que marcan el día a día de cualquier estudiante: **el tiempo que dedicamos a estudiar** y **el uso de TikTok o Instagram justo antes de dormir**.

**Objetivo:** analizar con datos reales si el uso de redes sociales por la noche es una práctica común entre los jóvenes, compararlo con información de organismos oficiales y reflexionar de forma ética sobre nuestro bienestar digital.

---

#### 📋 METODOLOGÍA
- **Instrumento:** formulario de Google Forms con 3 preguntas (edad, horas de estudio y uso de redes antes de dormir).
- **Muestra:** 50 personas de entre **[mín]** y **[máx]** años, con una edad promedio de **[promedio]** años.
- **Procesamiento:** las respuestas se organizaron en Excel (columnas = campos, filas = registros), se eliminaron duplicados y celdas vacías, y todos los resultados se calcularon con fórmulas automáticas (`CONTAR.SI`, `PROMEDIO`, `MAX`, `MIN`, `CONTARA`).
- **Ética y protección de datos:** la encuesta fue 100 % anónima. No se pidieron nombres, correos ni teléfonos; cada participante se identifica solo con un folio (R01–R50). Los datos se resguardan en una carpeta organizada de Google Drive y se usan únicamente con fines educativos.

---

#### 📊 RESULTADOS

**[Imagen: Gráfico 1]**

**Gráfico 1. Uso de TikTok o Instagram antes de dormir.**
Como se observa en el Gráfico 1, el **[X]%** de los encuestados (**[n]** de 50) usa TikTok o Instagram antes de dormir, mientras que solo el **[Y]%** no lo hace. Esto indica que revisar redes sociales en la cama es un hábito **[mayoritario / minoritario]** en nuestra muestra.

**[Imagen: Gráfico 2]**

**Gráfico 2. Tiempo de estudio diario.**
En el Gráfico 2 vemos que el **[X]%** de los participantes estudia más de 2 horas al día y el **[Y]%** estudia menos. **[Elige una: "La mayoría dedica un tiempo considerable al estudio" / "Más de la mitad estudia menos de 2 horas diarias, lo que sugiere un área de oportunidad"]**.

**Cruce de variables (Tabla 4 del Excel).**
De las **[n]** personas que usan redes antes de dormir, **[a]** estudian más de 2 horas y **[b]** no. Entre quienes **no** usan redes por la noche, **[c]** estudian más de 2 horas. *(Si quieres compararlo en porcentajes: a ÷ n × 100.)* Con 50 personas no podemos afirmar que un hábito cause el otro, pero sí observar si aparecen juntos.

---

#### 🌍 CONTRASTE CON DATOS ABIERTOS

Nuestros resultados no son un caso aislado; coinciden con lo que reportan organismos e investigaciones serias:

1. **UNESCO – Informe de Seguimiento de la Educación en el Mundo 2023, *"Tecnología en la educación: ¿una herramienta en los términos de quién?"*.** El informe advierte que el uso excesivo del celular se relaciona con un menor rendimiento académico y que la sola cercanía del teléfono puede distraer a los estudiantes. Esto concuerda con lo que observamos: un hábito digital tan extendido como el uso nocturno de redes compite directamente con el tiempo y la concentración que dedicamos al estudio.
   Fuente: https://www.unesco.org/gem-report/es/technology

2. **Investigación científica sobre pantallas y sueño (Hale y Guan, 2015, *Sleep Medicine Reviews*).** Esta revisión sistemática analizó 67 estudios sobre niños y adolescentes y encontró que, en alrededor del 90 % de ellos, el uso de pantallas se asociaba con dormir más tarde y durante menos tiempo. Este comportamiento coincide con nuestro hallazgo de que el **[X]%** de los encuestados usa redes justo antes de dormir.
   Fuente: https://doi.org/10.1016/j.smrv.2014.07.007

3. **Recomendación de sueño para adolescentes (Academia Estadounidense de Medicina del Sueño, AASM).** Recomienda que los jóvenes de 13 a 18 años duerman **de 8 a 10 horas** diarias. Si las redes retrasan la hora de dormir, es más difícil cumplir esa meta.
   Fuente: https://aasm.org/resources/pdf/pediatricsleepdurationconsensus.pdf

4. **México – INEGI, ENDUTIH 2023.** La Encuesta Nacional sobre Disponibilidad y Uso de Tecnologías de la Información en los Hogares reporta que la gran mayoría de la población mexicana de 6 años o más usa internet, principalmente desde el celular, y que el acceso a redes sociales es una de sus actividades más comunes.
   Fuente: https://www.inegi.org.mx/programas/dutih/2023/

> 📝 Antes de entregar, abre cada enlace y comprueba que funcione. Si quieres, agrega una cifra textual de la fuente que más te guste.

---

#### ⚖️ CONCLUSIÓN ÉTICA

Los datos de esta investigación muestran que el **[X]%** de los jóvenes encuestados usa TikTok o Instagram antes de dormir, una tendencia que coincide con lo que señalan la UNESCO y la investigación científica sobre el sueño. Esto no significa que las redes sociales sean "malas", sino que **su uso sin límites puede afectar dos recursos valiosos para un estudiante: el descanso y la concentración**.

Como ciudadanos digitales, tenemos la responsabilidad de usar la tecnología de forma **consciente y equilibrada**. Algunas acciones concretas:
- Dejar el celular fuera de la cama o activar el modo "No molestar" 30 a 60 minutos antes de dormir.
- Usar las herramientas de bienestar digital del teléfono para ver y limitar el tiempo en cada aplicación.
- Proteger nuestro tiempo de estudio con bloques sin notificaciones.

Este proyecto también demostró que **es posible investigar con ética**: recopilamos información útil sin exponer la identidad de nadie, respetando la privacidad de los 50 participantes. Proteger los datos personales de otros es tan importante como cuidar nuestro propio bienestar digital.

---

## Paso 5 – Entrega · **TÚ**
1. Google Sites → **Publicar** → elige una dirección web → en "Quién puede ver mi sitio" debe aparecer **Público**. Copia el enlace.
2. Abre el enlace en una **ventana de incógnito** para confirmar que se ve sin iniciar sesión.
3. Sube a la plataforma escolar:
   - `Apellido_Nombre_Proyecto.xlsx` (con tu apellido y tu nombre)
   - El enlace público de tu Diario Online

## ✅ Checklist según la rúbrica
- [ ] Carpeta `Proyecto_Competencias_Digitales` pública y con subcarpetas (5 pts)
- [ ] Archivo nombrado `Apellido_Nombre_Proyecto.xlsx`
- [ ] Formulario sin nombres, correos ni teléfonos
- [ ] 50 filas, sin duplicados ni vacíos y con "✔ Base completa" (10 pts)
- [ ] `CONTAR.SI`, `SUMA`, `PROMEDIO`, `MAX` y `MIN` con referencias de celda (15 pts)
- [ ] Sitio con menú e íconos, gráficos nítidos, análisis, contraste con fuentes y conclusión ética (20 pts)
