# DSIM Changelog

All notable changes to the Warehouse Sign-In System (DSIM) are documented here for end users.

---

## January 2026 Update

### 🎉 New Features

#### Admin Dashboard - Settings Tab
A brand new **Settings** tab has been added to the Admin Dashboard where you can customize your experience:

- **🔊 Sound Controls**
  - Enable/disable notification sounds
  - Choose different sounds for fresh logins vs EPJ swaps
  - Adjust notification volume with a slider
  - Test sounds before saving with dedicated test buttons

- **👁️ Visual Notifications**
  - Toggle visual popup notifications on/off
  - Keep sounds but hide popups if you prefer

- **⏱️ Auto-Refresh Options**
  - Choose how often the dashboard updates (30s, 1min, 2min, 5min, or manual only)
  - Reduces unnecessary refreshes if you prefer manual control

- **📊 Compact Mode**
  - Enable compact tables for denser data display
  - Useful when you need to see more rows at once

All settings are saved to your browser and persist between sessions.

---

#### Live Checkouts - Type Column & "Worked" Checkbox
- **Type Column**: Now shows whether a checkout is a "Fresh" login (blue) or "EPJ Swap" (orange)
- **Different Sounds**: Fresh logins play an ascending chime, swaps play a double-beep
- **Worked Checkbox**: Admins can mark drivers as "worked" - row highlights green when checked
  - Use this to track which drivers have already been processed/assisted
  - Status persists even after page refresh

---

#### Better Empty States
When tables have no data, you'll now see friendly messages explaining why and what to do:
- "No Active Checkouts" - All drivers are checked in
- "No EPJs Configured" - EPJ Status sheet needs data
- "No Matching Users" - Adjust your search filters
- "No Users Found" - No one has registered yet

---

#### Helpful Tooltips
Hover over buttons to see what they do:
- Admin Checkout → "Manually check out a driver"
- Refresh → "Refresh checkout list now"
- Edit/Reset/Delete user buttons all have descriptive tooltips
- Export CSV, Select All, and more

---

### 🚀 Performance Improvements

#### Faster Search
- Search boxes now use "debouncing" - the app waits 300ms after you stop typing before searching
- Prevents lag when typing quickly in User Management or Load Support Dashboard

#### Smarter Caching
- The system now only clears relevant cached data instead of everything
- Logging out only clears your session, not all EPJ data
- This means faster page loads and less waiting

#### Keyboard Navigation
- EPJ cards in the Admin Dashboard can now be navigated with Tab
- Press Enter or Space on a focused card to open its context menu
- Improves accessibility for keyboard users

#### Mobile-Friendly Touch Targets
- Buttons and inputs are now larger on mobile devices (minimum 44px height)
- Easier to tap without accidentally hitting the wrong thing

---

### 🔧 Bug Fixes & Improvements

- **Swap Detection**: More reliable identification of EPJ swaps vs fresh checkouts
- **Sound System**: Fixed test buttons in Settings to properly play sounds
- **Session Handling**: Better automatic recovery when sessions expire
- **Error Handling**: API calls now retry automatically if they fail (up to 3 times)

---

## How to Access New Features

1. **Settings Tab**: In Admin Dashboard, click the "⚙️ Settings" tab in the navigation
2. **Worked Checkbox**: Visible in the Live Checkouts tab for each driver row
3. **Type Column**: Automatically shows in Live Checkouts - no action needed
4. **Tooltips**: Just hover over any button to see the hint

---

## Tips for Best Experience

- **Adjust Volume**: If sounds are too loud/quiet, use the Settings slider
- **Test Sounds First**: Use the test buttons to preview before waiting for a real login
- **Use Compact Mode**: If you manage many drivers, enable compact tables
- **Reduce Auto-Refresh**: Set to 2+ minutes if you don't need real-time updates

---

*For technical documentation, see the [Admin Dashboard User Guide](Admin-Dashboard-User-Guide.md).*
