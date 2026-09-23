# PPTX Academic Figure — Patterns

Generic layout recipes. Adapt module names/content to the paper; do not copy project-specific copy.

## Common pipeline skeleton

```
[ Input ] → [ Module A ] → [ Module B ] → [ Module C ] → [ Output ]
                 └──────── [ Eval / Verify bar ] ────────┘
```

- Side columns: tall and narrow
- Center modules: near-square cards in one row
- Optional bottom bar under center modules only

## Module header pattern

1. Bold title: `Module name`
2. Optional short role, only when useful or requested
3. Schematic content
4. Optional short footer caption

## Image slots

| Role | Typical AR | Notes |
|------|------------|-------|
| Main subject / result | Source-dependent | Normalize visible scale when comparisons require it; preserve AR |
| Object thumbnail / sample | ≈ 1.0 | Square grid when suitable |
| Side column frame | ≈ 0.45 | Tall strip |

For paired source/result rows, use equal square viewports when requested, with proportional content inside. Use proportional padding for isolated objects or a suitable crop for continuous content. Do not stretch a narrow subject to fill its square.

Grey placeholder = image slot only. Text chips and legends are not grey photo blocks.

## Accent usage

- Pastel fill for module identity
- Stronger accent on small UI (weight chip, verify arrow, bar chart)
- Keep most of the canvas greyscale

## Do

- Nest frames; mute palette; Times titles; consistent gaps
- Match photo AR before inserting real images
- Inspect a targeted preview after a batch of edits; use a full view for composition changes

## Don't

- Purple-glow / AI-slide look, emoji, thick stacked shadows
- Huge sans titles in tiny cards
- Scale geometry without scaling fonts
- Replace valid existing connectors without a task-specific reason
- Add a redundant slide title unless requested
