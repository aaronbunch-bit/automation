# ReflexAI weekly scheduling capacity

Weekly assigned scheduling allows no more than three simultaneous booked ReflexAI blocks per audience (`Sales Group` / weekly `Supergroup`), across all simulation names. It checks the full time interval. A block ending at 10:20 does not overlap one starting at 10:20.

The existing `WEEKLY_MAX_REPS_PER_SIM_SLOT` setting remains compatible. Values above three are capped at three; valid lower limits remain supported. Each booked offer counts once, even when it lists several simulations. Active booked offers are reloaded before committing under a script lock so competing weekly runs cannot use the same available capacity.

Booked windows must have an end time or a valid recorded duration. Missing audience or invalid timing stops scheduling rather than silently excluding a booking from the count. Existing bookings are not moved or canceled.

## Validation

Run `node --test tests/weekly-capacity.test.cjs`.

## Deployment boundary

Update `TrainingScheduler_Weekly_Assigned_Mode.gs` in the existing Apps Script project to activate this change. GitHub changes alone do not update that project. This repository does not contain the implementation of `tsCommitTrainingToAssembled_` or the full offer-based booking flow. This safeguard covers the weekly assigned path and persisted active bookings in `TS Offers`; manual Assembled bookings and other writers must use the same audience overlap check and lock to guarantee a system-wide limit. Verify those paths in the deployed project before claiming full coverage.
