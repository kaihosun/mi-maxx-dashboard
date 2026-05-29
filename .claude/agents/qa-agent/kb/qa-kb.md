# QA KB — MI Technologies MTY MAXX

## Quality Standard
ISO 9001:2015 — referenced across all project SOPs

## Quality Policy
`10. Documentos del Sistema/1. Politica de Calidad/POL-MTY-QA-001.md`

## SOP Locations
| Area | SOP File |
|------|---------|
| Incoming | `1. Incoming/3. SOP/SOP-MTY-INC-001.md` |
| Logistics | `2. Logistica/3. SOP/SOP-MTY-LOG-001.md` |
| Sorting | `3. Sorting/3. SOP/SOP-MTY-SOR-001.md` |
| Duplicates | `10. Documentos del Sistema/SOP-MTY-DUP-001.md` |

## SOP Standard Sections (ISO 9001)
Each SOP must contain:
- Section 10: Control y Verificación / Checkpoints
- Section 11: Análisis de Riesgos (Clause 6.1)
- Section 12: Indicadores de Desempeño / KPIs (Clause 9.1)
- Section 13: Capacitación y Competencia (Clause 7.2)
- Section 14: Registros Generados (Clause 7.5.3)
- Section 15: Control de Cambios

## Code Quality Checklist
- [ ] No hardcoded department names (use SHEET_NAMES)
- [ ] All new TRANSLATIONS keys added for both Spanish and English
- [ ] UTF-8/UTF-8-BOM encoding handled on CSV reads
- [ ] No None values written to Excel cells
- [ ] Build script runs without errors
- [ ] Output file sheet count matches len(SHEET_NAMES) = 15
- [ ] All activity labels (actividad 1-21) present in TRANSLATIONS

## Data Audit Checklist
- [ ] CSV has header row with expected column names
- [ ] No empty rows in required columns (NOMBRE, actividad 1..N)
- [ ] No broken encoding characters (replacement char: `�`)
- [ ] Enriched CSVs (`_enriquecido.csv`) have superset of base CSV columns
- [ ] Week identifier (Semana) present in all rows

## Verdict Format
```
QA VERDICT: PASS / FAIL
Findings:
  [P] Item passes
  [F] Item fails — <specific issue> — <file>:<line>
Recommendation: <action required>
```

## Skills Available
- `quality-audit` — full pre-release quality check
- `verify-document` — validates a single SOP or process document against ISO 9001 structure
