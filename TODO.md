# Geofence Radius Increase to 3km Task

# Remove Complete GPS Features Task

## Status: In Progress

### Previous (Radius Increase):
- Radius changed to 3km ✅

### New Steps:
- [ ] Code.gs: Remove GPS validation, dist calc/enforce in markAttendance()
- [ ] index.html: Remove getLocation(), hide loc-pill, bypass GPS flows
- [ ] Verify no GPS prompt/errors
- [ ] Update TODO.md
- [ ] Complete

**Changes Summary:**  
Increase geofence radius from 1km (1000m) to 3km (3000m) in backend (Code.gs) and frontend (index.html) for consistency.
