# Changelog

## February 1, 2026

### New Features

#### 1. Dynamic Zone Loading for Force Check-In (Admin Dashboard)
- Force check-in modal now loads locations dynamically from the **Zones** sheet
- Replaced hardcoded location options with `getZonesList()` function call
- Zones are fetched when the modal opens, ensuring up-to-date options

#### 2. Overspill Driver EPJ Pickup Enhancement
- Overspill drivers (those without an EPJ) can now **pick up an EPJ mid-trip**
- New UI section: "Need an EPJ for Your Trip?" with:
  - Zone filter dropdown to filter available EPJs by location
  - EPJ selection dropdown showing only available EPJs
  - EPJ info display showing location and status of selected EPJ
  - Optional fields to update Route, Tractor, and Trailer numbers
  - "Get EPJ & Continue Trip" button
- New backend function: `overspillGetEpj()` to process EPJ pickup
- Trip details are updated and logged when driver picks up an EPJ

#### 3. EPJ Swap Modal Enhancements
- Added **zone filter** to EPJ swap modal for easier EPJ selection
- Added **EPJ info display** showing location and status of selected EPJ
- Drivers can now see where an EPJ is located before selecting it
- Filter persists while browsing available EPJs

### Bug Fixes & Performance Improvements

#### 4. Check-In Page Zone Loading Fix
- **Issue**: Zones were loading very slowly or not at all when drivers tried to check in
- **Root Cause**: The `getCheckinFormData()` function was making slow async calls
- **Solution**: 
  - Zones are now **preloaded server-side** directly into the HTML template
  - Removed dependency on client-side async loading for zones
  - Check-in button is now enabled immediately on page load
  - Added fallback "Unknown Location" option if zones fail to load

#### 5. getZoneOptions() Improvements
- Added comprehensive error handling with try/catch
- Changed from `A1:A` (entire column) to `getLastRow()` for more efficient data retrieval
- Added logging for debugging zone loading issues
- Returns fallback option if sheet is missing or empty
- Cache validation to check for empty cached values

#### 6. getCheckinFormData() Optimization
- Simplified function to return early when no EPJ needs to be checked
- Removed slow direct sheet lookup for store-only status
- Uses cached EPJ statuses only (avoids slow spreadsheet reads)
- Added overall try/catch with logging

#### 7. Event Listener Safety
- Added null checks for swap modal event listeners
- Prevents JavaScript errors when elements don't exist in certain views

### Code Changes Summary

| File | Changes |
|------|---------|
| `code.js` | Added `getZonesList()`, `overspillGetEpj()`, optimized `getCheckinFormData()`, improved `getZoneOptions()`, updated `getUserView()` to preload zones |
| `CheckInForm.html` | Added Overspill EPJ pickup UI, enhanced EPJ swap modal with zone filter, preloaded zones via template scriptlet, added console logging for debugging |
| `AdminDashboard.html` | Updated force check-in modal to load zones dynamically |

### Technical Notes

- Zones are cached for 6 hours (21600 seconds) via `CacheService`
- Session tokens are cached for 24 hours (86400 seconds)
- EPJ statuses use script-level caching for performance
- Template uses `<?!= zoneOptions ?>` scriptlet for server-side zone injection
