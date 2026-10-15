# Contributing to TTS Word

Thanks for your interest! This is a small personal project and isn't under active development, but issues and pull requests are welcome and will be reviewed when time allows.

## Reporting bugs and ideas

Open an [issue](../../issues). For bugs, include your Word and Windows versions, whether you were online (edge-tts) or offline (Piper), and the helper log at `%LOCALAPPDATA%\TTS Word\helper.log` if you have it installed.

## Making changes

1. Fork the repo and create a branch from `main`.
2. Set up the helper and the add-in as described in [Development](README.md#development).
3. Make your change and test it in Word with `npm start`.
4. Before opening the pull request, run these in `addin/`:

   ```powershell
   npm run lint
   npx tsc --noEmit -p . --types office-js --skipLibCheck
   npm run build
   ```

   CI runs the same checks on every pull request.

## Guidelines

- **Keep it simple for the end user.** The target user is comfortable in Word and not much beyond it: everything should work from the task pane, with no consoles, manual configuration or technical error messages.
- The UI text is in Spanish. Code comments may be in Spanish or English.
- Don't change the document while reading. Words are marked with `range.select()`, never by changing formatting, so the user's document and undo history stay untouched.
- [CLAUDE.md](CLAUDE.md) has detailed architecture notes (in Spanish), including some known lint false positives.
- Keep pull requests focused: one fix or feature per pull request.

By contributing, you agree that your contributions are licensed under the [MIT License](LICENSE).
