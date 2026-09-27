# Roadmap

## Completed — 0.8.0

- Template-driven DOCX analysis and formatting core
- Safe output that never overwrites either input file
- Persistent asynchronous task storage
- Spring Boot upload, status, result, and report APIs
- Vue 3 + Element Plus upload and result page
- One-command full application build and startup
- Real-file verification with the included template and manual
- Rule and structure-issue confirmation page before document modification
- Cover, bilingual abstracts, Word TOC, references, acknowledgements, and appendix formatting
- Template-specific figure captions inside borderless tables, including image/caption page binding
- Roman front-matter and Arabic body page numbering with Word field refresh
- Dual template modes: real samples/styles or prose-only formatting requirements
- Source-first structure recognition for prose-only templates, with the detected mode shown before confirmation
- Direct TOC paragraph spacing, cover auxiliary layout tables, figure containers, and body table cell formatting
- Abstract label indentation correction and conservative bibliography whitespace repair

## Next

1. Add a database-backed task and template repository.
2. Expand the validation engine to report every remaining equation, floating text-box, cross-reference, and complex-field mismatch.
3. Add optional AI-assisted parsing for ambiguous natural-language requirements that the deterministic parser cannot resolve.
4. Add user accounts, quotas, orders, and payment only after core formatting accuracy is stable.
