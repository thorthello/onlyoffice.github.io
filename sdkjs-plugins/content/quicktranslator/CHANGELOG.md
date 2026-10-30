# Changelog

All notable changes to the QuickTranslator plugin are documented in this file.

## [1.3.3] - 2026-10-30
### Added
- Dynamic initial target language detection via `window.Asc.plugin.info.lang`, adapting to the user's ONLYOFFICE UI language with English default fallback.

### Fixed
- Fixed false positive network block detection in `isHtmlResponse()` by eliminating word matching on translated text ("Access denied", "autenticazione", etc.).
- Localized all internal status, error, and fallback notifications to pure English.

### Changed
- Removed DeepL translation engine and API key UI due to browser CORS restrictions on client-side requests, streamlining QuickTranslator into a 100% free, zero-config tool.
- Updated Privacy and Data Disclosure in `README.md` to accurately disclose Google Translate, LibreTranslate, and MyMemory.

## [1.3.2] - 2026-09-15
### Added
- Official ONLYOFFICE Marketplace store metadata, screenshots, and assets (`resources/store/`).
- 3 high-resolution store showcase screenshots (`screen_1.png`, `screen_2.png`, `screen_3.png`).
- Full Dark Theme support with dynamic `onThemeChanged` event listener and `@media (prefers-color-scheme: dark)`.
- Vendor attribution (`"offered": "Mattia Musiello"`).
- `translations/langs.json` marketplace language manifest.
- Pure English interface conforming to ONLYOFFICE marketplace guidelines.

### Fixed
- Fixed rapid selection race condition on double and triple clicks using intelligent debouncing and AbortController request cancellation.

## [1.3.1] - 2026-09-14
### Added
- Local and CDN fallback bridge for `plugins.js`.
- Automatic cursor position monitoring (`onTargetPositionChanged`).

## [1.0.0] - 2026-09-01
### Initial Release
- Real-time text translation across 33+ languages.
- Automatic source language detection.
- One-click document insertion for Docs, Sheets, and Slides.
