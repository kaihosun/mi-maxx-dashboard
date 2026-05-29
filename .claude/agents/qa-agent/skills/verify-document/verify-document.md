# Skill: Verify Document

## Purpose
Validate a single SOP, work instruction, or process document against the ISO 9001:2015 structure expected in this project.

## When to Use
- After updating or creating any SOP in the project folders
- When onboarding a new department's documentation
- When a user asks "is this SOP complete?"

## Steps
1. Read the target document (`.md` or `.docx` converted to `.md`)
2. Check for all required ISO 9001 sections:
   - [ ] Section 1: Objective / Objetivo
   - [ ] Section 2: Scope / Alcance
   - [ ] Section 3: Definitions / Definiciones
   - [ ] Section 4: Responsibilities / Responsabilidades
   - [ ] Section 5: Procedure / Procedimiento
   - [ ] Section 10: Control Points / Puntos de Control y Verificación
   - [ ] Section 11: Risk Analysis (Clause 6.1)
   - [ ] Section 12: KPIs / Indicadores de Desempeño (Clause 9.1)
   - [ ] Section 13: Training & Competence (Clause 7.2)
   - [ ] Section 14: Records Generated (Clause 7.5.3)
   - [ ] Section 15: Change Control
   - [ ] Approvals section
3. Check document metadata: document code (e.g., `SOP-MTY-INC-001`), revision date, version
4. For bilingual documents: verify both Spanish and English versions have matching section counts
5. Cross-reference: if SOP references other documents, verify those files exist in the project

## Verdict Output
```
DOCUMENT REVIEW: <filename>
Standard: ISO 9001:2015
Language: ES / EN / Bilingual

Section Checklist:
  [✓/✗] Section <N>: <name>

Metadata:
  Code: <doc code>
  Version: <version>
  Date: <date>

Missing: <list any missing sections>
Overall: COMPLETE / INCOMPLETE

Recommendation: <action>
```
