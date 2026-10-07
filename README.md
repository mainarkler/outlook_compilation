# Outlook Compilation

VBA prototype for Outlook 2019 + Word 2019.

Current specification:
- Source: Outlook folder configured in modConfig.bas.
- Subject must contain a date in the form: Вопрос на DD/MM/YYYY.
- User runs CompileQuestions and enters the target date.
- A Word template on a shared drive is copied; the template itself is never modified.
- The first 6 Word paragraphs of the template are preserved.
- Questions are inserted starting at paragraph 7.
- Each question is rendered as: Вопрос № N / Вопрос выносится <sender display name> / email body text excluding tables.
- Tables found in email bodies are appended at the end under ПРИЛОЖЕНИЯ.
- Each table is labelled Приложение к вопросу № N.
- Word automation is late-bound, so no Word VBA reference is required.

This repository contains VBA source files, not an Outlook OTM binary. Import the BAS modules into Outlook VBA (Alt+F11). See INSTALL.md.