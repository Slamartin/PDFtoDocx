# Tasks: PDFtoWord launch

Plan: `tasks/plan.md`. Spec: `SPEC.md`. Check boxes as tasks complete.

## Task 1: Commit SPEC.md and the plan, mark the spec approved

**Description:** Put the approved spec and this plan under version control so later sessions and CI
have them. Change the spec's status line from draft to approved with the date.

**Acceptance criteria:**
- [x] `SPEC.md` status line reads "approved 2026-10-12"
- [x] `SPEC.md`, `tasks/plan.md`, `tasks/todo.md` are committed on `main`

**Verification:**
- [x] `git log --oneline -1` shows the commit; `git status` clean

**Dependencies:** None

**Files likely touched:** `SPEC.md`, `tasks/plan.md`, `tasks/todo.md`

**Estimated scope:** XS

---

## Task 2: Add ruff, fix findings, run it in CI

**Description:** Add ruff as a dev dependency with the spec's configuration, format the code, fix lint
findings without changing behaviour, and run both checks in the CI `test` job. Two commits: format first,
lint fixes second, so the behaviour-neutral reflow is isolated. CI edit is ask-first per the spec; the
spec approval covers it, but say so in the commit.

**Acceptance criteria:**
- [x] `pyproject.toml` has `ruff` in the dev group and `[tool.ruff]` with `line-length = 120` and an
      explicit `select = ["E4", "E7", "E9", "F", "I"]` (ruff 0.17's wider defaults flag the converters'
      intentional broad excepts, so the classic set is pinned instead of `extend-select`)
- [x] `uv run ruff check .` and `uv run ruff format --check .` exit 0 for every committed file
      (the uncommitted `tests/test_pptx_internals.py` from another session has one E741 left)
- [x] `.github/workflows/ci.yml` runs both checks before pytest
- [x] `uv run pytest -q` still passes (170 tests with the other session's new file present)

**Verification:**
- [x] Tests pass: `uv run pytest -q`
- [x] Lint clean: `uv run ruff check . && uv run ruff format --check .`
- [x] Manual check: `git diff --stat` of the format commit touches no test assertions

**Dependencies:** Task 1

**Files likely touched:** `pyproject.toml`, `uv.lock`, `.github/workflows/ci.yml`, formatted files under
`app/` and `tests/`

**Estimated scope:** M

---

## Task 3: Replace the placeholder repository URL

**Description:** Put the real GitHub URL `https://github.com/Slamartin/PDFtoDocx` (from Martin,
2026-10-12) into every place that carried the old `martinslama/pdftoword` placeholder.

**Acceptance criteria:**
- [x] `grep -r "martinslama/pdftoword"` returns nothing
- [x] `app/config.py` default, `fly.toml` `[env]`, `pyproject.toml` `[project.urls]`, and the
      `index.html` footer fallback all show the real URL
- [x] `GET /api/limits` returns the real URL (asserted in `test_client_side_limit_endpoint`)

**Verification:**
- [x] Tests pass: `uv run pytest -q tests/test_api_limits.py -k limit_endpoint`
- [x] Manual check: `/api/limits` and the footer `href` both carry the repo URL

**Dependencies:** Task 1 (needs the repo name decision)

**Files likely touched:** `app/config.py`, `fly.toml`, `pyproject.toml`, `app/static/index.html`

**Estimated scope:** S

---

## Checkpoint 1: Clean local repo
- [x] `uv run pytest -q` passes
- [x] `uv run ruff check .` and `uv run ruff format --check .` clean
- [x] Decide: keep charging a rate-limit token on 413 (current) or reorder the middlewares
      -> keep (Martin, 2026-10-12)
- [x] Review with Martin before publishing

---

## Task 4: Create the GitHub repo, push `main`, get CI green (manual)

**Description:** Martin creates a public repository on GitHub (browser or `gh repo create`). The agent
adds the remote, pushes `main`, and watches the first CI run. The `docker` job is the first time the
image builds on Linux CI; the `deploy` job will fail until Task 9 and that is expected.

**Acceptance criteria:**
- [x] `git remote -v` shows `origin` on GitHub; `main` pushed (2026-10-12)
- [x] CI `test` job green; CI `docker` job green (container converts `sample.pdf`), run 38444213982;
      `deploy` failed as expected (no `FLY_API_TOKEN` yet)
- [x] Repository visibility is public; GitHub shows the AGPL-3.0 licence badge (2026-10-13). Before the
      switch the history was rewritten to the GitHub noreply commit email and the old CI runs deleted.

**Verification:**
- [x] Manual check: GitHub Actions page shows green for `test` and `docker`
- [x] Manual check: the footer URL `https://github.com/Slamartin/PDFtoDocx` returns 200 without login

**Dependencies:** Task 3

**Files likely touched:** none (git remote only)

**Estimated scope:** XS

---

## Checkpoint 2: Public source
- [x] CI green for `test` and `docker`
- [x] Source link works; AGPL obligation met

---

## Task 5: Install flyctl, create the Fly app, first deploy (manual)

Done 2026-10-13 with flyctl v0.4.116. Instead of `fly launch` (interactive, may rewrite `fly.toml`):
`fly apps create pdftoword --org personal`, `fly ips allocate-v6` and `fly ips allocate-v4 --shared` (avoids
the paid dedicated IPv4), then `fly deploy --remote-only --ha=false` (one machine, so no `fly scale count 1`).
Fly asked for billing before it would create the app. Remote image 314 MB.

**Description:** Install flyctl (`winget install flyctl` or the PowerShell installer), `fly auth signup`
or `fly auth login`, then `fly launch --no-deploy --copy-config` (confirm app name and region, keep the
generated config identical to `fly.toml`), `fly deploy --remote-only`, `fly scale count 1`. If the app
name changes, update `fly.toml` and README.

**Acceptance criteria:**
- [x] `fly status` shows one machine in the chosen region, state started or stopped (machine
      83099ebddd7698, `ams`, shared-cpu-1x, 1024 MB, 512 MB swap, health check passing)
- [x] `https://pdftoword.fly.dev/healthz` returns `{"status":"ok"}` over HTTPS
- [x] `fly.toml` in the repo matches what Fly has (unchanged; machine config shows the same services,
      concurrency, auto stop/start and health check)

**Verification:**
- [x] Manual check: `curl https://pdftoword.fly.dev/healthz` (200); `/api/limits` shows the public
      profile and the CSP header is present
- [x] Manual check: `fly logs` shows "Application startup complete"; the only error is one health check
      that ran a second before uvicorn was listening, passing from the next check on

**Dependencies:** Task 4

**Files likely touched:** `fly.toml` (only if the name or region changes), `README.md`

**Estimated scope:** XS

---

## Task 6: Post-deploy functional checks

**Description:** Run the hosted checks from the spec against the public URL with curl and the browser:
guards, real files, retention, privacy page, headers.

**Acceptance criteria:**
- [x] 30 MB upload gets 413 quickly (0.08 s, 64 KB sent); the sixth quick upload gets 429 with
      `Retry-After: 23`. While this PC was at 429, a POST from the Fly machine itself passed the limiter,
      so the buckets key on the real client IP (`Fly-Client-IP`), not the proxy address.
- [ ] Both real decks convert and download (waiting for the deck files); `sample.pdf` (3.5 s) and
      `sample.pptx` convert and the DOCX files open; a junk `.pdf` named `Confidential-Payroll-2026.pdf` gets
      the generic 422 and the name appears in neither the response nor `fly logs`
- [x] `/privacy` and the footer links work; response headers include the CSP; `/docs`, `/redoc` and
      `/openapi.json` are 404; no cookies; the only absolute URL in the page is the GitHub source link
- [x] After `RESULT_TTL_S` the result link returns 404 (200 at 840 s, 404 at 960 s, machine kept awake and
      not restarted in between); `fly status` shows the machine still healthy

**Verification:**
- [x] Manual check: the curl sequence from the local verification, pointed at the public URL (2026-10-13)
- [x] Manual check: page shows "Hosted service" and the real limits line in the footer (headless Edge)

**Dependencies:** Task 5

**Files likely touched:** none

**Estimated scope:** XS

---

## Task 7: Measure the Fly proxy against a long silent conversion

**Description:** Generate an 80-page, image-heavy PDF with PyMuPDF into a temp folder (not committed),
upload it to the public URL, and time the response. The goal is a conversion that takes clearly more
than 60 s on the shared CPU. Record the outcome: completed after N s, or connection closed after N s
with the worker still running in `fly logs`.

**Acceptance criteria:**
- [x] A test PDF exists whose conversion takes > 60 s on Fly: a 22-page synthetic PDF (dense styled text,
      ruled table, noise image per page) runs into the 90 s worker timeout; 10 pages of it take 35 s
- [x] Result recorded in `SPEC.md` open question 1 with the measured number

**Verification:**
- [x] Manual check: `curl -w` gave `HTTP 504 in 90.9 s`: the proxy delivered the silent response
- [x] Manual check: `fly logs` shows `conversion failed ... error=timeout` at 18:57:41 and the 504 sent

**Dependencies:** Task 6

**Files likely touched:** `SPEC.md` (open questions)

**Estimated scope:** XS

---

## Task 8 (conditional): Adjust the timeout if Task 7 fails

Not needed (2026-10-13): the proxy delivered a 90.9 s response, so the 90 s timeout stays.

**Description:** Only if the proxy cut the response. Lower `CONVERT_TIMEOUT_S` in the Dockerfile to just
under the measured proxy limit and lower `MAX_PAGES`/`MAX_SLIDES` so typical files finish inside it. This
is an ask-first change (public limits). If the usable timeout is too short for real files, stop and plan
the `202` + `GET /api/jobs/{id}` polling API as its own spec instead.

**Acceptance criteria:**
- [ ] Martin approved the new numbers
- [ ] Dockerfile env, README table and SPEC.md updated together
- [ ] Redeployed; Task 7's PDF now gets a clean 504 or converts within the limit

**Verification:**
- [ ] Tests pass: `uv run pytest -q`
- [ ] Manual check: repeat Task 7 upload against the new deploy

**Dependencies:** Task 7

**Files likely touched:** `Dockerfile`, `README.md`, `SPEC.md`

**Estimated scope:** S

---

## Task 9: Automatic deploys from CI (manual secret)

**Description:** Create a Fly deploy token (`fly tokens create deploy -x 999999h`), add it as the
`FLY_API_TOKEN` repository secret on GitHub, push a trivial change to `main`, and confirm the `deploy`
job deploys.

**Acceptance criteria:**
- [ ] The `deploy` job is green on a push to `main`
- [ ] `fly releases` shows the CI-triggered release

**Verification:**
- [ ] Manual check: GitHub Actions `deploy` job log ends with a successful `flyctl deploy`

**Dependencies:** Task 5

**Files likely touched:** none

**Estimated scope:** XS

---

## Checkpoint 3: Live
- [ ] Task 6 checks all pass on the public URL
- [ ] Task 7 measured and recorded; Task 8 done or not needed
- [ ] Push to `main` deploys automatically

---

## Task 10: Record the live URL and decisions

**Description:** Put the public URL in the README's hosted section, resolve the spec's open questions
(URL, Fly name, 413 token decision, proxy result), and propose the vault update (research note hosting
section gets the URL and the proxy measurement) for Martin's OK.

**Acceptance criteria:**
- [ ] README shows the live URL and the final limits
- [ ] SPEC.md open questions reduced to the custom domain
- [ ] Vault proposal listed in the reply; written only after OK

**Verification:**
- [ ] Manual check: README limits match `Dockerfile` env and `GET /api/limits`

**Dependencies:** Task 9

**Files likely touched:** `README.md`, `SPEC.md`, vault `Research/pdftoword-converter.md`

**Estimated scope:** S

---

## Checkpoint: Complete
- [ ] All tasks checked
- [ ] Spec, README and vault agree with what is deployed
