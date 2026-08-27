Build a Python project called "Mock Email Generator" to create testing-friendly mock emails from 60 original Nomura trade emails.

Requirements:
1. Input: 60 emails (.eml/.msg) with optional Excel/PDF attachments.
2. Generate 5 mock versions per original = 300 mocks.
3. Preserve the original email body, HTML/plain-text formatting, tables, structure, and attachment layout as much as possible.
4. ONLY mask sensitive identity information:
   - Counterparty/company/entity names
   - Nomura entity names
   - Sender/From, To, CC email addresses
   - Personal names
   - Signature names/entity information
5. DO NOT modify trade data such as amount, currency, dates, trade IDs, quantities, prices, etc.
6. Do not create separate parsers for each email format. Build a generic detection/replacement engine that works regardless of headers such as C/P, Counterparty, Client, Entity, etc.
7. Excel: preserve workbook structure, sheets, formatting, formulas and layout; replace only detected sensitive values.
8. PDF: preserve the existing content/layout and replace only detected sensitive values where technically possible. Identify whether PDFs are text-based or scanned.
9. Use deterministic Python techniques first: regex, dictionaries, context rules and configurable mappings. Do not use AI/LLMs.
10. Create a discovery mode that scans all 60 emails/attachments and produces a reviewable list of detected names, entities, email addresses and signatures before masking.
11. Maintain consistent mapping: the same original person/entity/email must always map to the same mock value within a generated mock.
12. Generate 5 different synthetic mappings per original while keeping all non-sensitive trade data unchanged.
13. Add validation to ensure original sensitive values are not present in generated mocks.
14. Keep all rules/configuration separate from the code so new formats and entities can be added without changing the core engine.

First create the project structure, requirements.txt, configuration files and discovery/analyzer module. Do not generate the full 300 mocks until the detection results can be reviewed.
