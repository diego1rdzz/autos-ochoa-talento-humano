# Autos Ochoa · Desarrollo organizacional e innovación en talento humano

Reto Final de **Administración del Talento Humano** (Parte 2 del Sistema Integral de Gestión de Talento, SIGT).
Informe estratégico de 14 páginas con ocho herramientas de desarrollo organizacional para el plan de expansión
«Ochoa 5» de Autos Ochoa (Monterrey, N.L.): de 1 a 5 sucursales, de 8 a 34 colaboradores y de 5 a 60 unidades
vendidas al mes. Todos los módulos se anclan al perfil crítico (**Gerente de Sucursal**) y a los objetivos
estratégicos OE-1, OE-2 y OE-3 definidos en la Parte 1.

**Entregable:** [`entregables/Autos_Ochoa_Reto_Final_Desarrollo_Organizacional.pdf`](entregables/Autos_Ochoa_Reto_Final_Desarrollo_Organizacional.pdf)

## Contenido del informe

| Página | Sección | Producto |
|---|---|---|
| 1 | Portada | — |
| 2 | Resumen ejecutivo | Matriz de alineación módulo → restricción → OE → KPI → inversión |
| 3 | Módulo 1 · Capacitación y Desarrollo | Plan de Formación Blended «Academia Ochoa»: objetivo, audiencia, temas, metodología presencial/digital y KPI Kirkpatrick |
| 4–5 | Módulo 2 · Clima Laboral y Experiencia del Colaborador | Cuestionario «Pulso Ochoa» (24 reactivos + eNPS), resultados simulados y plan de mejora a 6 meses |
| 6 | Módulo 3 · Employee Branding y Fidelización | EVP, plan de comunicación interno y externo, estrategias de fidelización |
| 7–8 | Módulo 4 · Gestión del Desempeño y Competencias | Modelo de 6 competencias (2 técnicas, 2 de liderazgo, 2 digitales/I4.0) y evaluación 360° |
| 9 | Módulo 5 · Responsabilidad Social y Talento | Tres políticas ASG: equidad salarial, inclusión en selección, balance vida-trabajo |
| 10 | Módulo 6 · Sucesión y Talento Clave | 9-Box con Colaboradores A, B y C y mapa de sucesión |
| 11 | Módulo 7 · Digitalización y Metodologías Ágiles | Herramientas HR Tech y Scrum/Kanban de RH |
| 12 | Módulo 8 · Innovación Organizacional y Excelencia | «Copiloto Modelo Ochoa» + Kaizen y certificación recomendada |
| 13 | Tablero integral de KPIs y hoja de ruta a 12 meses | — |
| 14 | Conclusiones, riesgos de implementación y fuentes | — |

## Regenerar el PDF

El informe se escribe en `informe/informe.html` y se imprime a PDF tamaño carta con Chromium (Playwright).

```bash
npm install
npm run build        # genera entregables/Autos_Ochoa_Reto_Final_Desarrollo_Organizacional.pdf
```

El script falla si el contenido de alguna página desborda su área útil. Variables opcionales:
`PREVIEW_DIR=/ruta` guarda un PNG por página y `CHROMIUM_PATH=/ruta/a/chromium` usa un Chromium ya instalado.

Las tipografías (Inter y Source Serif 4, licencia SIL Open Font) se incluyen en `informe/fonts/`.
