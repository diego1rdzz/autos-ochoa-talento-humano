# Autos Ochoa · Plan de talento 2026–2027

Reto Final de **Administración del Talento Humano** (Parte 2 del Sistema Integral de Gestión de Talento, SIGT).
Documento de 15 páginas, escrito en lenguaje sencillo para explicar al equipo qué se va a hacer internamente, por qué,
quién lo hace, cuándo y cómo sabremos que funcionó. Son ocho acciones de desarrollo organizacional para el plan de
expansión «Ochoa 5» de Autos Ochoa (Monterrey, N.L.): de 1 a 5 sucursales, de 8 a 34 personas y de 5 a 60 autos al mes.
Todo se ancla al perfil crítico (**Gerente de Sucursal**) y a los tres objetivos de la Parte 1:
OE-1 gerente listo antes de abrir, OE-2 sucursal rentable en 4 meses y OE-3 que el talento se quede.

**Entregable:** [`entregables/Autos_Ochoa_Reto_Final_Desarrollo_Organizacional.pdf`](entregables/Autos_Ochoa_Reto_Final_Desarrollo_Organizacional.pdf)

## Contenido del documento

| Página | Sección | Qué incluye |
|---|---|---|
| 1 | Portada | — |
| 2 | En pocas palabras | Por qué ahora, los tres objetivos y las ocho acciones con su meta y su costo |
| 3 | Módulo 1 · Capacitación y Desarrollo | Academia Ochoa: objetivo, a quién va dirigida, temas, formación presencial y en línea, y cómo se mide |
| 4–5 | Módulo 2 · Clima Laboral y Experiencia | Encuesta «Pulso Ochoa» (27 preguntas), resultados simulados y plan de mejora de 6 meses |
| 6 | Módulo 3 · Marca Empleadora y Permanencia | Propuesta de valor (EVP), plan de comunicación interno y externo, y acciones para retener |
| 7–8 | Módulo 4 · Desempeño y Competencias | Seis competencias (2 técnicas, 2 de liderazgo, 2 digitales) y evaluación 360° |
| 9 | Módulo 5 · Responsabilidad Social | Tres políticas: mismo puesto mismo sueldo, contratar por lo que sabes hacer y tiempo para tu vida |
| 10 | Módulo 6 · Sucesión y Talento Clave | Mapa de talento 9-Box con los Colaboradores A, B y C y plan de sucesión |
| 11 | Módulo 7 · Digitalización y Agilidad | Cuatro herramientas digitales y trabajo en ciclos de 2 semanas (Scrum/Kanban) |
| 12 | Módulo 8 · Innovación y Excelencia | Asistente con las prácticas del fundador y certificación recomendada |
| 13 | Cómo mediremos y cuándo | Tablero de indicadores y calendario de 12 meses |
| 14 | Qué cambia para cada quien | Qué cambia para gerentes, asesores, oficinas y familia; próximos 90 días y preguntas frecuentes |
| 15 | Conclusiones | Conclusiones, riesgos y fuentes |

## Regenerar el PDF

El informe se escribe en `informe/informe.html` y se imprime a PDF tamaño carta con Chromium (Playwright).

```bash
npm install
npm run build        # genera entregables/Autos_Ochoa_Reto_Final_Desarrollo_Organizacional.pdf
```

El script falla si el contenido de alguna página desborda su área útil. Variables opcionales:
`PREVIEW_DIR=/ruta` guarda un PNG por página y `CHROMIUM_PATH=/ruta/a/chromium` usa un Chromium ya instalado.

Las tipografías (Inter y Source Serif 4, licencia SIL Open Font) se incluyen en `informe/fonts/`.
