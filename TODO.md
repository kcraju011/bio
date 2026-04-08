# Geofence Radius Increase to 3km Task

## Status: ✅ COMPLETE

### Steps from Approved Plan:
- [x] Gather info from files
- [x] Create detailed edit plan and get approval
- [x] Create this TODO.md
- [x] Edit Code.gs: Update FENCE_RADIUS_M = 3000 and comment
- [x] Edit index.html: Update FENCE_M = 3000
- [x] Verify changes with read_file
- [x] Update TODO.md with completion

**Changes Summary:**  
Geofence radius increased from 1km (1000m) to 3km (3000m):  
- Code.gs: `FENCE_RADIUS_M = 3000; // 3km geofence radius`  
- index.html: `FENCE_M = 3000; // 3km to match backend`  

Backend now allows attendance from up to 3km from SIT coordinates (13.3260801, 77.1261350). Frontend displays updated distance feedback.

**Changes Summary:**  
Increase geofence radius from 1km (1000m) to 3km (3000m) in backend (Code.gs) and frontend (index.html) for consistency.
