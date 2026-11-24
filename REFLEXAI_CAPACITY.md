# ReflexAI audience overlap protection

Applied to Apps Script on November 24, 2026. Weekly scheduling allows at most three simultaneous blocks per audience across simulation names and start times. Adjacent blocks are allowed. The targeted scheduler retains its existing stricter two-person limit.

## Active booking paths

- Weekly assigned scheduling: interval-based selection in `TrainingScheduler_Weekly_Assigned_Mode.gs`, followed by the shared final gate in `ReflexAI_Capacity_Guard.gs` through the existing Assembled writer.
- Independent targeted scheduling: `targeted/Canada_Targeted_Scheduler.gs` calls the same gate through the `ReflexCapacity` library, pinned to version 11 of the main project. Its existing forecast, roster, duration, two-person limit, preview integrity, and publish confirmation checks remain in place.
- Book It is retired, per Aaron. No Book It deployment was created, refreshed, or reactivated. Remaining old source is not treated as an active route. Other files in the targeted project are removal placeholders.

## Shared gate

The main project and library callers use the same project script lock and reservation ledger. The gate re-reads live Assembled activities, resolves audiences from Assembled people/queues, and checks peak simultaneous occupancy immediately before writing. It conservatively counts all Training activities in the audience, consistent with the targeted scheduler. The hard maximum is three; callers may impose a lower limit.

Reservations are persisted before the API write. Uncertain or not-yet-visible writes continue to occupy capacity until the block's end; this prevents delayed API reads or subsequent spreadsheet failures from allowing a fourth block. Visible reservations are not counted twice. Failed reads, invalid intervals, unknown candidate audiences, and unclassified existing ReflexAI blocks prevent writes. Writes are never automatically retried.

A pending/uncertain reservation may require review after a failed API call. Existing scheduled blocks are not rescheduled or canceled. Direct manual edits and unrelated external automation cannot be prevented by Apps Script; blocks already visible in Assembled are included at the time of the final check.

## Live projects

- Main: `1SkbzTG2l1NuKfCFoSN42REOcstJ9J1t4_LEq-3lS_giyXCgJgX3Bj45i`
- Targeted: `1CoSH8g3EsSQX64vwCzUlfSaMOagYh3uYmeS_kio3bDyvf_u6Pkn_tcRl`
- Shared library version: 11. Library-only release; no web-app release.

The targeted manifest keeps its existing spreadsheet-current-only and external-request scopes. No sharing changes or new credentials were added.

## Validation

`node --test tests/*.test.cjs` passes 13 tests covering overlapping and adjacent intervals, mixed simulations, separate audiences, saved-booking loading, invalid data, final rechecking, lock ownership, read failures, uncertain writes, and reservation deduplication.

Live read-only validation at 3:56 PM CT resolved 544 Assembled people and upcoming Training activities across all five audiences with zero bookings written. The targeted project's live library test at 3:59 PM CT confirmed that library version 11 rejects invalid input before any external read/write. Saved editor contents were compared with the tested files. No live booking or messaging function was executed for testing.

`ReflexAI_Workshop_Scheduling.gs` now contains the full live module; the previous GitHub copy was truncated. Preserve the source and the library dependency when making future changes.
