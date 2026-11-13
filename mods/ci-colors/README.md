# ci-colors

A copy of the [prismantis](https://github.com/NahumLitvin/prismantis) mod (v0.46.0, MIT, by Nahum Litvin) with its own colors, width and text styles. It redraws Claude's replies: tables, headings, code, mermaid diagrams, numbers, paths and links.

Turn prismantis off before you turn this on. Both draw the same replies, so only one should run.

## Install

At the prompt of a terminal session:

```
/plugin install ci-colors --marketplace scottbrowncpa-creator/ci-public
```

From a shell:

```
claude plugin install ci-colors --marketplace scottbrowncpa-creator/ci-public
```

In the desktop app: click + next to the prompt box, then Plugins, Add marketplace with `scottbrowncpa-creator/ci-public`, then Add plugin.

## Change the look

Everything you would edit is in `hooks/settings.ts`, one setting per line with a comment:

- `MAX_WIDTH` - widest a reply, table or code block draws, in columns. 0 uses the full window.
- `HEADING_STYLE`, `BOLD_TEXT`, `ITALIC_TEXT`, `UNDERLINE_LINKS` - how text is styled.
- `BRAND` - the colors. `USE_BRAND = false` goes back to the base theme.

Font face and size are not mod settings. Set them in the app you read in (the terminal's profile settings, or the desktop app's appearance settings).

Commands are the same as prismantis, under `/ci-colors` (for example `/ci-colors setup`, `/ci-colors demo`).

## What changed from prismantis

- Name, command, settings keys and state renamed to `ci-colors`, so it cannot collide with prismantis.
- `hooks/settings.ts` added; `hooks/theme.ts`, `hooks/render.tsx` and `hooks/register.tsx` read it.
- The issue hint is off by default.
- `hooks/brand.test.tsx` checks colors, width and text styles. Run `claude plugin test .` in this folder.

## License

[MIT](LICENSE), by Nahum Litvin. Bundled third-party code and palettes are credited in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
