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

- [x] Define shortage score formula:
  - shortage_score = gap_z * abs(corr)
  - sorted by shortage_score (replaces gap_z sort)
- [x] Define improvement threshold:
  - IMPROVEMENT_DELTA = 0.20 (shortage_score drop >= 0.20 = "improved")
  - STABLE_DELTA = 0.05 (change < 0.05 = "stable")
- [x] Define confidence level:
  - n >= 60 → 高, n >= 30 → 中, n < 30 → ⚠低
  - shown as "信頼度" column in shortage table

## Priority 3: Weekly Tracking

- [x] Add weekly comparison view for one short_id (load_my_history / show_weekly_tracking_section).
- [x] Show trend for top 3 shortage features (plotly line chart).
- [x] Add weekly action plan output:
  - keep (shortage_score <= 0)
  - improve (delta >= IMPROVEMENT_DELTA)
  - watch (delta <= -IMPROVEMENT_DELTA)
  - stable (|delta| < STABLE_DELTA)

## Notes

- Character-specific factors remain out of scope.
- Play-volume metrics are advisory-only (separate from core skill score).
- Use short_id as the primary key for personal diagnosis.
