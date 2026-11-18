# S4 - Manto v6 (2026)

Copia del informe "S4 - Manto v5_2" (el original no fue modificado), como plantilla `.pbit`.

## Cómo usarla
1. Abre `S4 - Manto v6 (2026).pbit` con Power BI Desktop.
2. Si pide credenciales para `10.5.0.52 / POWERBIS4`, usa las mismas de siempre.
3. Espera a que cargue los datos y **Guardar como** un `.pbix` nuevo, con otro nombre.
4. Publica solo ese archivo nuevo, nunca encima del informe original.

## Fuentes (dbo, servidor 10.5.0.52, base POWERBIS4)
- Tabla `FACT_PS_PRESUP_PEP` (presupuesto)
- Vista `VW_ZOSTB_EJECPEP` (ejecutado)
- Vista `VW_ZOSTB_COMPRPEP` (comprometido)

## Cambios respecto al original
- Power Query `FACT_PS_PRESUP_PEP`: filtro `Periodo = "2025"` -> `"2026"`
- Tabla DAX `comprometido`: `Año_cargo = "2025"` -> `"2026"`
- Tabla DAX `calendario`: llega hasta el 31/12/2026 (antes 31/12/2025)
- Informe: 33 filtros/segmentaciones de año 2025 -> 2026 y el texto "*Año 2025 a Junio" -> "*Año 2026 a Junio"
- Se quitó el vínculo con el informe publicado en el servicio, para que esta copia no pueda sobrescribirlo.

Todo lo demás (páginas, visuales, medidas, relaciones, tablas auxiliares) es idéntico.
