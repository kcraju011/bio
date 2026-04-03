# BioAttend Fix: TypeError output.addHeader is not a function

## Approved Plan Steps
- [x] **Step 1**: Edit `code.gs` - Fix `doOptions()` function (replace 3 `addHeader()` → `setHeaders()`) ✅
- [x] **Step 2**: Edit `code.gs` - Fix `jsonOut()` function (replace `addHeader()` → `setHeader()`) ✅
- [ ] **Step 3**: Test in Google Apps Script editor (run `doOptions()`, check no errors)
- [ ] **Step 4**: Deploy new web app version, update frontend API URL if needed
- [ ] **Step 5**: Test full flow: OPTIONS preflight + frontend API calls (Sign In/Register)

**Status**: Starting edits...

