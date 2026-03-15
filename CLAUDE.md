# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the project

```bash
python main.py
```

## Git workflow

After every meaningful change, commit and push to GitHub so work is never lost and can be reverted at any point:

```bash
git add <files>
git commit -m "concise description of what changed and why"
git push
```

- Commit frequently — after each feature, fix, or logical unit of work
- Keep commit messages clean and specific (not "update" or "fix stuff")
- Always push after committing; the remote at https://github.com/Dustinf98/claude-code-test is the source of truth
