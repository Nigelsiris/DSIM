# Changelog

## February 7, 2026

### Performance Improvements

#### Batch Spreadsheet Operations
- **New `adminBatchMaintenanceStatus()`**: Processes multiple EPJ status changes in a single API call
- **10x faster bulk operations**: Uses `setValues()` instead of individual `appendRow()` calls
- **Less server load**: Single transaction for all selected EPJs

#### Lazy Loading for Tabs
- **Fleet Management tab**: Content loads only when tab is first accessed
- **Maintenance tab**: EPJ grid loads on-demand
- **Reports tab**: Data fetched when tab is viewed
- **Announcements tab**: Loads independently of initial page load
- **Result**: Faster initial page load, better perceived performance

### New Features

#### Dashboard Metrics Cards
- **5 real-time metrics**: Available, Checked Out, Maintenance, Total EPJs, Utilization %
- **Auto-updating**: Metrics refresh every minute with live data
- **Color-coded**: Green (available), Blue (checked out), Orange (maintenance)
- **Utilization percentage**: Shows what percentage of fleet is actively in use

#### Browser Notification System
- **Push notifications**: Get alerts even when Admin Dashboard tab is in the background
- **Permission handling**: Graceful request for notification permission
- **Per-user preferences**: Enable/disable via Settings tab
- **Notification types**: Different styling for checkouts vs swaps
- **Test button**: Verify notifications are working
- **Automatic integration**: Hooks into existing checkout/swap detection

### Major Feature: Maintenance Mode Overhaul

Complete redesign of the Maintenance tab in the Admin Dashboard for improved usability.

#### New Visual EPJ Grid
- **Click-to-select interface**: EPJs displayed as cards that can be clicked to select
- **Visual status indicators**: 
  - Orange border/background for EPJs in maintenance
  - Green border for available EPJs
  - Red border (grayed out) for checked-out EPJs
- **Grouped sections**: EPJs automatically sorted into "In Maintenance", "Available", and "Checked Out" sections
- **Checkbox multi-select**: Each card has a checkbox for quick selection

#### Bulk Actions
- **Select All Visible**: Select all EPJs matching the current filter
- **Clear Selection**: Deselect all EPJs at once
- **Put in Maintenance**: Bulk action to take multiple EPJs out of service
- **Return to Service**: Bulk action to return multiple EPJs to active duty
- **Filter dropdown**: Filter view by status (All, In Maintenance, Available, Checked Out)

#### Maintenance Action Modal
- **Reason field**: Required when putting EPJs in maintenance
- **Resolution notes field**: Track what was fixed when returning EPJs to service
- **Visual EPJ list**: Shows all selected EPJs before confirming action
- **Real-time feedback**: Success/error messages during bulk operations

#### Enhanced Maintenance History
- **EPJ filter**: Filter history by specific EPJ
- **Event type filter**: Filter by "Started Maintenance" or "Returned to Service"  
- **Date range filters**: Filter by start and end dates
- **Clear filters button**: Reset all filters at once
- **Resolution notes column**: Now displays what was fixed

#### Backend Changes
- `adminSetMaintenanceStatus()`: Updated to accept resolution notes
- `getMaintenanceHistory()`: New function to fetch maintenance log with filtering support

---

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
