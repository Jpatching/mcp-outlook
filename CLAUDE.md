# CLAUDE.md — Developer & AI Collaboration Guidelines

Guidance for developers and AI agents working on `mcp-outlook`.

---

## 🔒 Non-Negotiable Invariants

1. **Human-in-the-Loop Safety**: Never bypass draft mode. All outbound messages MUST default to the Outlook `Drafts` folder unless explicit `confirm_send=True` is provided.
2. **Prompt Injection Defense**: All untrusted email body content MUST remain wrapped in `<untrusted_email_body_content>` XML tags.
3. **Zero Secrets**: Never commit real user email addresses, tenant IDs, OAuth tokens, or passwords.
4. **Local MAPI Focus**: Keep the zero-cloud Windows COM MAPI architecture first-class.

---

## 🛠️ Development & Testing Workflow

Always run tests before committing:

```bash
# Run test suite
pytest -v

# Run single test module
pytest tests/test_security.py -v
```

---

## 🌿 Git & Branching Strategy

- **`main` is production**: Must always be stable, tested, and deployable.
- **Feature Branches**: Use descriptive branch names:
  - `feat/<feature-name>` for new functionality
  - `fix/<bug-name>` for bug fixes
  - `test/<test-name>` for test additions
  - `docs/<doc-name>` for documentation updates
- **Conventional Commits**: Format commit messages as:
  - `feat: ...`, `fix: ...`, `test: ...`, `security: ...`, `docs: ...`, `chore: ...`
