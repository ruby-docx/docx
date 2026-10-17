---
name: tdd-development
description: The test-driven workflow for ALL code changes in ruby-docx/docx — features and bug fixes alike, not just bugs. Use whenever implementing, fixing, or changing behavior in this repo. Drives the full loop: sync master, write a failing test first (confirm red), implement to green, then branch → PR → CI → merge → cleanup. Invoke before starting any code change here.
---

# Test-driven development workflow (docx)

Apply this to **every** behavior change in this repo — new features, bug fixes,
refactors that change behavior. Always write the test first and watch it fail
before writing the implementation.

## Environment (read first)

- Local Ruby is 4.0.x where **`bundle exec rake spec` is broken** (`ostruct` is
  no longer a default gem). **Run tests with `bundle exec rspec spec`** directly,
  or a single file/example with `-e`.
- Output is noisy with rubygems warnings. Filter them for readable results:
  `bundle exec rspec spec 2>&1 | grep -vE 'warning:|previous definition'`.

## The loop

1. **Sync master and branch.**
   ```sh
   git switch master && git fetch origin master && git merge --ff-only origin/master
   git switch -c <topic-branch>
   ```
2. **Write a failing test first.** Add a spec that reproduces the desired
   behavior or the bug, in the spec file of the class that owns the behavior
   (see "Where tests go" below). Name a regression test after the issue, e.g. a
   comment `# Regression test for #NNN`.
3. **Confirm RED.** Run just that example and verify it fails for the expected
   reason:
   ```sh
   bundle exec rspec spec/docx/containers/paragraph_spec.rb -e "<example text>" 2>&1 | grep -vE 'warning:|previous definition'
   ```
4. **Implement the minimal change** to make it pass. Keep it focused.
5. **Confirm GREEN** for the new test and **run the full suite** to catch
   regressions:
   ```sh
   bundle exec rspec spec 2>&1 | grep -vE 'warning:|previous definition' | tail -4
   ```
6. **Commit, push, open a PR.** End the commit body with the Co-Authored-By
   trailer. Write a PR body explaining the bug/feature, the fix, and the tests;
   add `Closes #NNN` when it resolves an issue.
7. **Check CI, then merge.**
   ```sh
   gh pr checks <pr> --repo ruby-docx/docx
   gh pr merge <pr> --repo ruby-docx/docx --merge   # merge commit, matching repo history
   ```
8. **Update master and clean up.**
   ```sh
   git switch master && git fetch origin master && git merge --ff-only origin/master
   git branch -D <topic-branch> && git push origin --delete <topic-branch>
   ```

## Where tests go

Specs mirror `lib/` — one spec file per class:

- **Unit specs** (`spec/docx/containers/*_spec.rb`, `spec/docx/elements/*_spec.rb`):
  build the node from an XML fragment with `w_node('<w:p>…</w:p>')` /
  `w_body(…)` (`spec/support/xml_helpers.rb`), pass `document_properties` as a
  plain hash, and stand in for the Document with `instance_double(Docx::Document, …)`.
  Prefer this for element behavior — no .docx fixture needed.
- **`spec/docx/document_spec.rb`**: only what Docx::Document itself owns —
  opening the package, loading parts (styles, rels, headers, footers),
  document-wide lookups, and save/stream/replace_entry.
- **`spec/integration/*_spec.rb`**: end-to-end scenarios on real .docx fixtures —
  save-and-reopen round trips, or bugs that only reproduce with a real file's
  structure (Office365, WPS, Google Docs, …).

Use `fixture_path('x.docx')` for fixtures and `tmp_path('out.docx')` for saved
files (a per-example temp dir; never write into `spec/fixtures`).

## Conventions specific to this repo

- **Test fixtures (.docx) can be crafted programmatically.** Read the entries of
  an existing fixture with rubyzip, mutate `word/document.xml` / `word/styles.xml`
  / `word/_rels/document.xml.rels` with Nokogiri, and rewrite the zip. Used to
  build split-placeholder, malformed-rels, missing-styles, and bookmark fixtures.
  Run such generators with `bundle exec ruby` and commit the resulting binary.
- **Keep new constructor arguments optional** (`def initialize(node, props = {},
  doc = nil)`) so existing direct instantiations and callers keep working
  (backward compatibility — e.g. threading `document_properties` through Table →
  Row → Cell → Paragraph).
- **Be honest about scope.** When a report (e.g. a fuzzer issue) spans several
  causes, separate gem-side logic bugs (which you fix) from raw exceptions thrown
  by underlying libraries (rubyzip/Nokogiri) on genuinely corrupt input (out of
  scope). Say so in the PR rather than implying a total fix.
- **Don't state something is done before it is.** Order operations so any comment
  or claim is true when posted (e.g. add a `Co-authored-by` trailer to the commit
  *before* telling a contributor they are credited).

## After a batch of merges

A series of fixes/features usually warrants a release — see the `release-docx`
skill (confirm the version with the user; a new public method is a minor bump).
