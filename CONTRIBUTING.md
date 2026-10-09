# Contributing

Thanks for wanting to make this better! Bug reports, ideas and pull requests are all welcome.

## Reporting a bug or suggesting an idea

Open an issue using one of the templates. For bugs, the most useful things are:

- What you did (e.g. "shared a Facebook event link from the iPhone Shortcut")
- What you expected, and what happened instead
- The message you saw, or the row that ended up in the sheet (remove anything private first)

Security problems: please use **Security → Report a vulnerability** on GitHub instead of a public issue.

## Working on the code

1. Fork the repo and clone your fork.
2. Set up your own test copy with clasp (see "For developers" in the [README](README.md)). Please test against your own sheet, not someone else's.
3. Make your change, `clasp push`, and try it in the sheet and on the deployed page (**Deploy → Test deployments** gives you a `/dev` link that always runs your latest code).
4. Open a pull request describing what changed and how you tested it. Screenshots help for anything visual.

## Guidelines

- **Keep it no-server.** Everything runs in the owner's Google account. Anything that needs a separate server or paid service (beyond the Claude API) is probably a separate project.
- **Never commit secrets or personal data:** API keys, Shortcut keys, `.clasp.json`, real people's events or fliers. Use made-up examples.
- **Treat everything from the public as untrusted:** form fields, shared links, flier text and images. Text written to the sheet goes through `cell_()`. Images go through `validImage_()`. Functions only the owner should run call `ownerOnly_()`. Internal helpers end in `_` so the page can't call them.
- **Accessibility matters:** real buttons and labels, visible focus, readable contrast, works with a keyboard and screen reader.
- **Match the existing style:** plain JavaScript (Apps Script V8 on the server, vanilla JS in the page), no build step, comments only where the *why* isn't obvious.

## Code of conduct

Be kind. This is a hobby project for sharing fun things to do with your friends.
