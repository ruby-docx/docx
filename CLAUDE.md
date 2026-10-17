# docx — project guide for Claude

`docx` is a Ruby gem (ruby-docx/docx) for reading and editing `.docx` files: a
thin wrapper over rubyzip + Nokogiri.

## Repo & access

- `origin` points directly at the upstream repo `github.com/ruby-docx/docx`
  (not a fork). The user is a **maintainer with push access**, so branches and
  tags can be pushed to `origin` and PRs merged directly.

## Environment gotchas

- Local Ruby is 4.0.x where **`bundle exec rake spec` fails** (`ostruct` is no
  longer a default gem). **Run tests with `bundle exec rspec spec`** instead
  (or a single example with `-e`).
- Test output is noisy with rubygems warnings; filter for readability:
  `... 2>&1 | grep -vE 'warning:|previous definition'`.

## Conventions

- **Always work test-first.** Every behavior change (feature or bug fix) follows
  the `tdd-development` skill: write a failing test, confirm red, implement,
  confirm green + full suite, then branch → PR → CI → merge → cleanup.
- **Specs mirror `lib/`, one file per class.** Unit-test elements from XML
  fragments (`w_node`), keep `document_spec.rb` to what Document owns, and put
  real-file round trips in `spec/integration/` (details in `tdd-development`).
- **Keep new constructor arguments optional** (`def initialize(node, props = {},
  doc = nil)`) to preserve backward compatibility for existing callers.
- **Be honest about scope.** When a report spans multiple causes, separate
  gem-side logic bugs (fix them) from raw exceptions thrown by rubyzip/Nokogiri
  on genuinely corrupt input (out of scope) — and say which is which.
- **Don't claim something is done before it is.** Order operations so any comment
  or status is true at the moment it's posted.

## Releasing & contributor credit

- Releases are tag-driven; use the `release-docx` skill. Confirm the version with
  the user first (RubyGems publish is irreversible; a new public method is a
  **minor** bump). Do **not** edit `CHANGELOG.md` (stale since v0.7.0; the GitHub
  auto-generated notes are the source of truth).
- Auto-generated release notes credit each PR by its title and its opener (the
  maintainer). To get an **original contributor's** handle into the notes, embed
  the handle in the **PR title** (e.g. "... (original work by @user in #NN)") and
  add a `Co-authored-by:` trailer — co-authors alone do not appear in the notes.

## Current state & roadmap (as of v0.13.0, 2026-05-31)

- Shipped: header/footer **reading** (`#headers`/`#footers`), **write-back** on
  save, and **bookmark scanning in headers/footers**. All bug-labeled issues are
  resolved; the only open issues are question/enhancement (#157, #91).
- Not yet done (good next work):
  - Resolve the **main document part name via `[Content Types].xml`** instead of
    globbing `word/document*.xml`; `#update` currently writes back to a
    hardcoded `word/document.xml`, so Office 365 `document2.xml` is not updated.
  - **Numbering** support (`word/numbering.xml`) — requested in #73.

## Skills

- `release-docx` — cut a new gem release (version bump → tag → CI publishes).
- `tdd-development` — the test-first workflow for any code change here.
