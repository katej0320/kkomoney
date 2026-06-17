# kkomoney Commit Message Convention

## Format

```
<type>(<scope>): <subject> (#issue)

<body (optional)>
```

## Type

| Type | Description |
|---|---|
| `feat` | New feature |
| `fix` | Bug fix |
| `refactor` | Code change that neither fixes a bug nor adds a feature |
| `style` | Changes that don't affect code meaning (formatting, CSS, etc.) |
| `docs` | Documentation only changes (README, etc.) |
| `chore` | Build process, config, dependency, or other maintenance changes |
| `perf` | Performance improvement |
| `test` | Adding or fixing tests |

## Scope (optional)

Indicates the area affected, e.g. `transaction`, `budget`, `auth`, `home`, `ui`.

## Subject rules

- Start with lowercase
- No trailing period
- Use imperative mood (`add`, not `added` or `adds`)
- Keep under 50 characters when possible

## Issue reference

If the change relates to a tracked bug/feature number, append it at the end:

```
(#5)
```

## Breaking changes

If the change breaks existing behavior, add a line in the body starting with `BREAKING CHANGE:`.

## Examples

```
fix(transaction): reset moneyType when switching from transfer (#5)

fix(auth): add password reset flow (#7)

feat(home): filter banner by current month (#9)

refactor(app): split App.js into separate components

BREAKING CHANGE: moves transaction state into context, components importing useState directly need update
```
