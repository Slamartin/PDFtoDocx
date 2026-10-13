# Implementation Plan: PDFtoWord launch

## Overview

The hardened code, Docker image and CI are in the repo (3 commits on `main`, 26 tests, container
verified locally on 2026-10-11). This plan covers what is left between that state and a live, verified
hosted service per `SPEC.md`: lint tooling the spec mandates, the placeholder URL, publishing the repo,
the first Fly deploy, verifying the one unknown (Fly proxy vs a silent 90 s response), and closing the
spec's open questions. Several steps need Martin's accounts (GitHub, Fly) and are marked **manual**.

## Architecture Decisions

- No new code paths unless the proxy test fails. The only code tasks are ruff fixes and the URL constant.
- Ruff: `line-length = 120`, default rules plus `I` (import order); format with `ruff format`. Added as a
  dev dependency through `uv add --dev`, run in the CI `test` job. The spec approved this; the CI edit
  is still called out because CI changes are ask-first.
- Fly: one `shared-cpu-1x` 1 GB machine, scale-to-zero, remote build. Keep `CONVERT_TIMEOUT_S=90` until
  Task 7 measures the proxy. If the proxy cuts the response, prefer lowering the timeout over building the
  polling API now; the polling API is a separate plan.
- Proxy test file: an 80-page image-heavy PDF generated with PyMuPDF (page cap is 80), so the conversion
  runs well past 60 s on the shared CPU. Generated into a temp folder, not committed.

## Task List

### Phase 1: Repo hygiene (local, no accounts needed)
- [x] Task 1: Commit SPEC.md and the plan, mark the spec approved
- [x] Task 2: Add ruff, fix findings, run it in CI
- [x] Task 3: Replace the placeholder repository URL

### Checkpoint 1: Clean local repo
- [x] `uv run pytest -q` passes, `uv run ruff check .` and `uv run ruff format --check .` are clean
- [x] All Phase 1 work committed on `main` (another session's converter work stays uncommitted)
- [x] Review with Martin before publishing

### Phase 2: Publish (manual: GitHub account)
- [x] Task 4: Create the GitHub repo, push `main`, get CI green

### Checkpoint 2: Public source
- [x] CI `test` and `docker` jobs green on GitHub
- [x] Repo is public, LICENSE shows AGPL-3.0 on the GitHub page
- [x] Footer link on the page opens the repo

### Phase 3: Deploy (manual: Fly account)
- [x] Task 5: Install flyctl, create the Fly app, first deploy
- [ ] Task 6: Post-deploy functional checks
- [x] Task 7: Measure the Fly proxy against a long silent conversion
- [x] Task 8 (conditional): Adjust the timeout if Task 7 fails (not needed)
- [x] Task 9: Automatic deploys from CI

### Checkpoint 3: Live
- [ ] All Task 6 checks pass against the public URL
- [x] Task 7 result recorded in SPEC.md open questions
- [x] A push to `main` deploys without manual steps

### Phase 4: Close out
- [ ] Task 10: Record the live URL and decisions in README, SPEC.md and the vault

### Checkpoint: Complete
- [ ] SPEC.md has no open questions left except the custom domain
- [ ] Vault research note updated (after Martin's OK)

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Fly proxy closes a silent response before 90 s | High: long conversions fail for users | Task 7 measures it on day one with a generated 80-page PDF; Task 8 lowers `CONVERT_TIMEOUT_S` and the page caps, or schedules the polling API |
| Fly app name `pdftoword` already taken | Low | `fly launch` prompts for a name; update `fly.toml` and README |
| Ruff rewrites many lines in `pptx_to_docx.py` (1070 lines) | Medium: noisy diff hides real changes | Run `ruff format` in its own commit before `ruff check --fix`; tests must stay green between the two |
| Cold start after scale-to-zero (image is 1.39 GB) | Medium: first request waits several seconds | Accept for the free tier; note it on the privacy page if it exceeds ~10 s; `min_machines_running=1` is an ask-first change |
| Shared IPv4 on Fly, cost of always-on | Low | Scale-to-zero keeps it near zero; check `fly dashboard` billing after a week |
| GitHub Actions deploy job runs before secrets exist | Low | The job is gated on `main` push; it fails harmlessly until `FLY_API_TOKEN` is set in Task 9 |

## Open Questions

- Resolved: the repository is `Slamartin/PDFtoDocx` (Task 3 done 2026-10-12).
- Fly app name and region for Task 5 (current: `pdftoword`, `ams`).
- Resolved 2026-10-12: a 413 (oversize upload) keeps costing a rate-limit token (abuse-resistant default).
- Custom domain or `*.fly.dev`: not blocking, decide after launch.
