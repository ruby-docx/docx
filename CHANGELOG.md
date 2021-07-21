# Changelog

## v0.6.2

### Bug fixes

- Fix `Docx::Document#to_s` fails when given file has `document22.xml.rels` [#112](https://github.com/ruby-docx/docx/pull/112), [#106](https://github.com/ruby-docx/docx/pull/106)

## v0.6.1

### Bug fixes

- Use `Zip::File#glob` to match any `document.xml` [#104](https://github.com/ruby-docx/docx/pull/104)

### Chores

- Enable Coverall's coverage report [#102](https://github.com/ruby-docx/docx/pull/102)
- Add table write example to README.md [#99](https://github.com/ruby-docx/docx/pull/99)
- Replace Travis CI build with GitHub Action [#98](https://github.com/ruby-docx/docx/pull/98)
- Add ruby 3.0 to versions for testing on Travis CI [#97](https://github.com/ruby-docx/docx/pull/97)

## v0.6.0

### Enhancements

- Added support for hyperlinks (implemented [#70](https://github.com/ruby-docx/docx/pull/70) again) by ollieh-m and gopeter [#92](https://github.com/ruby-docx/docx/pull/92)

### Chores

- Drop ruby 2.4 from supporeted versions by satoryu [#93](https://github.com/ruby-docx/docx/pull/93)
- Refactoring `spec_helper` by satoryu [#90](https://github.com/ruby-docx/docx/pull/90)
- Starts measuring code coverage with coveralls by satoryu [#88](https://github.com/ruby-docx/docx/pull/88)

## v0.5.0

### Enhancements

- Added opening streams and outputting to a stream [#66](https://github.com/ruby-docx/docx/pull/66)
- Added supports for Office 365 files [#85](https://github.com/ruby-docx/docx/pull/85)

### Bug fixes

- `Docx::Document` handles a docx file without styles.xml [#81](https://github.com/ruby-docx/docx/pull/81)
- Fixes insert text before after were switched [#84](https://github.com/ruby-docx/docx/pull/84)

## v0.4.0

### Enhancements

- Implement substitute method on TextRun class. [#75](https://github.com/ruby-docx/docx/pull/75)

### Improvements

- Updates dependencies. [#72](https://github.com/ruby-docx/docx/pull/72), [#77](https://github.com/ruby-docx/docx/pull/77)
- Fix: #paragraphs grabs paragraphs in tables. [#76](https://github.com/ruby-docx/docx/pull/76)
- Updates supported ruby versions. [#78](https://github.com/ruby-docx/docx/pull/78)
