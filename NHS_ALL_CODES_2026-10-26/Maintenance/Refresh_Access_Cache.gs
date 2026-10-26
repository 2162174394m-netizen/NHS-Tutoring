/**
 * One-run helper after direct Access Registry edits.
 * Requires nhsa_invalidateAccessCache_() from NHS_Backend.gs.
 */
function refreshNHSAccessNow() {
  const revision = nhsa_invalidateAccessCache_();
  console.log("Access cache refreshed: " + revision);
  return revision;
}
