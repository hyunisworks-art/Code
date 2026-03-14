# Next Steps (SF6 Analysis Dashboard)

## Goal

- Improve readability of the dashboard.
- Define clear decision metrics for coaching output.

## Priority 1: Readability

- [ ] Reorder sections by user flow:
  1. sample size summary
  2. key findings
  3. personal coaching result
- [ ] Improve visual cues:
  - shortage = warm color
  - strength = cool color
  - add short legends for interpretation
- [ ] Reduce text length in auto summary:
  - 3-5 lines max
  - one-line takeaway at top

## Priority 2: Decision Metrics

- [ ] Define shortage score formula:
  - combine gap_z and abs(correlation)
  - keep scale simple and explainable
- [ ] Define improvement threshold:
  - what counts as "improved" week-to-week
  - what counts as "stable"
- [ ] Define confidence level:
  - based on target sample size and variance

## Priority 3: Weekly Tracking

- [ ] Add weekly comparison view for one short_id.
- [ ] Show trend for top 3 shortage features.
- [ ] Add weekly action plan output:
  - keep
  - improve
  - watch

## Notes

- Character-specific factors remain out of scope.
- Play-volume metrics are advisory-only (separate from core skill score).
- Use short_id as the primary key for personal diagnosis.
