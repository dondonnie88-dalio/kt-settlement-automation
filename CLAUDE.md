# kt-settlement-automation

Settlement/reconciliation automation scripts (`run_settlement.py`, `run_reconcile.py`, `make_sample_data.py`) plus Korean-labeled batch launchers (`run_매출분석.bat`, `run_정산자화화.bat`).

## Workflow for new automation/forecast features

Before building, work through these three steps in order (skipping straight to code causes rework):

1. **Problem definition** — one line: what pain point, what's the needed solution.
2. **PRD** — turn the problem definition into a requirements doc.
3. **Update this file** — once the PRD is settled, fold the relevant parts into CLAUDE.md so future sessions have the context.

Requirements change constantly on internal tooling. Keep the **core decision logic** stable and build the smallest MVP that covers it — don't over-build steps 4-10 of a 10-step feature before step 1-3 (the MVP) is validated, but keep those later steps in mind while designing so the MVP doesn't need to be torn up later.

## Three checkpoints for internal (non-developer-facing) automation

- **Accessibility** — favor data stores that are easy to inspect/edit/hand off (e.g. spreadsheets over a bespoke DB) when the data changes often and non-engineers need to touch it.
- **Extensibility** — even when only building the MVP slice, keep the full feature list in mind so later additions don't require a rewrite.
- **Security** — anything exposed beyond local/internal use needs at least basic auth (OAuth). Run a security review pass before deployment.

### Security review prompt (use before deploying anything externally reachable)

```
# Security-review the code below before deployment.

[paste code]
Stack: [framework] + [deploy env] + [DB/Auth stack]

# Assess these 15 items against the current code
01. CORS
02. CSRF
03. XSS + CSP
04. SSRF
05. AuthN/AuthZ
06. RBAC / tenancy
07. Least privilege
08. Input validation + SQLi
09. Rate limiting
10. Cookies / sessions
11. Secret rotation
12. HTTPS / security headers
13. Audit log
14. Error exposure
15. Dependency vulnerabilities

# Output format
| Item | Status | Risk | Fix direction |
Status: done / partial / none / n-a
Risk: High / Med / Low
```

Treat any "Med" or higher as something to actually fix before shipping.

## Notes

- Prefer the simplest adequate tool (a spreadsheet formula, Apps Script, n8n) over Claude Code when the task is genuinely simple; reach for Claude Code once project count/complexity grows past what those handle comfortably.
- Source: 요즘IT, "AI 도구 26개를 직접 만들며 알게 된 자동화 노하우" (클코나잇 시즌2 웨비나 정리), https://yozm.wishket.com/magazine/detail/3861/
