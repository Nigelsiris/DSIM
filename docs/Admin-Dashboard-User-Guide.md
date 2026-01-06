# Admin Dashboard User Guide

Welcome to the DSIM Admin Dashboard! This guide will help you manage EPJ equipment, monitor driver checkouts, and keep operations running smoothly.

---

## Getting Started

When you log in as an Admin, you'll see the **Admin Control Panel** with five main tabs at the top:

| Tab | What It's For |
|-----|---------------|
| **EPJ Status** | See all EPJs at a glance and quickly change their status |
| **Live Checkouts** | Monitor who currently has equipment checked out |
| **Users** | Add new drivers and manage user accounts |
| **Maintenance** | Put EPJs in/out of maintenance mode |
| **Fleet Management** | Add or remove EPJs from the system |

---

## EPJ Status Tab

### Understanding the EPJ Grid

Each EPJ is shown as a colored card:

- 🟢 **Green** = Available for checkout
- 🔴 **Red** = Currently checked out (shows driver name)
- 🟠 **Orange** = In maintenance

Cards marked with **[S]** are designated for **Store Delivery Only**.

### Quick Actions (Right-Click Menu)

Right-click any EPJ card to quickly:
- Set it to Available
- Set it to Checked Out
- Set it to Maintenance
- Toggle "Store Delivery Only" on/off

---

## Live Checkouts Tab

This tab shows you a real-time table of all current checkouts, including:

| Column | Information |
|--------|-------------|
| Driver | Who has the equipment |
| Carrier | The driver's carrier company |
| EPJ | Which EPJ they have (or "OS" for Overspill) |
| Type | **Fresh** (new checkout) or **Swap** (replaced another EPJ) |
| Truck | Truck number |
| Trailer | Trailer number |
| Route | Their assigned route |
| Time | When they checked out |
| Actions | Buttons to swap EPJ or force check-in |

### Understanding Checkout Types

- **Fresh** (blue badge) = Driver signed in and checked out equipment normally
- **Swap** (orange badge) = Driver's EPJ was swapped mid-shift (hover to see which EPJ they had before)

This helps you quickly identify drivers who may have experienced equipment issues.

### Features

- **Auto-Refresh**: The table updates every 30 seconds automatically
- **Manual Refresh**: Click the **⟳ Refresh** button anytime
- **Last Updated**: Shows when data was last refreshed

### Actions You Can Take

#### Swap EPJ
If a driver needs a different EPJ (equipment issue, wrong assignment, etc.):
1. Click **Swap** next to their checkout
2. Optionally enter why the old EPJ needs maintenance
3. Select the new EPJ from the dropdown
4. Click **Confirm Swap**

The system will automatically:
- Assign the new EPJ to the driver
- Put the old EPJ in maintenance (if selected)
- Log the change

#### Force Check-In
If a driver forgot to check in their EPJ:
1. Click **Check-In** next to their checkout
2. Confirm the action

This returns the EPJ to "Available" status.

#### Admin Checkout
Need to check out an EPJ on behalf of a driver?
1. Click **+ Admin Checkout**
2. Select the driver
3. Check "Overspill" if no EPJ is needed, OR select an EPJ
4. Enter truck, trailer, and route information
5. Click **Checkout EPJ**

---

## Users Tab

### Adding a Single User
1. Enter the **Username** (how they'll log in)
2. Enter a **Password**
3. Enter their **Carrier Name** (optional)
4. Select their **Role**: Driver, Admin, or Load Support
5. Click **Create User**

### Adding Multiple Users at Once
For bulk user creation:
1. Enter user data in the text box, one per line
2. Format: `username,password,Role,CarrierName`
3. Example:
   ```
   jsmith,password123,Driver,ABC Transport
   mjones,pass456,Driver,XYZ Logistics
   ```
4. Click **Create Multiple Users**

### Managing Existing Users
Click **Open User Management** to:
- Edit user roles and carriers
- Reset passwords
- Delete users

---

## Maintenance Tab

### Putting an EPJ in Maintenance
1. Select the EPJ from the dropdown
2. Enter a reason (optional but helpful for tracking)
3. Select **Put IN Maintenance**
4. Click **Update Status**

### Returning an EPJ from Maintenance
1. Select the EPJ
2. Select **Return FROM Maintenance**
3. Click **Update Status**

### Maintenance History
The table below shows recent maintenance events with:
- When it happened
- Which EPJ
- What event occurred
- Reason given
- Resolution notes

---

## Fleet Management Tab

### Adding a New EPJ
1. Enter the EPJ number (e.g., "EPJ-015")
2. Check **Store Delivery Only** if this EPJ should only be used for store runs
3. Click **Add EPJ**

### Removing an EPJ
1. Select the EPJ from the dropdown
2. Click **Remove EPJ**
3. Confirm the removal

> ⚠️ **Note**: You cannot remove an EPJ that is currently checked out. It must be checked in first.

---

## Notifications

### Driver Login Alerts
When a driver signs in, you'll:
- Hear an **ascending chime** (two tones going up)
- See a notification at the top of the screen
- See a badge showing recent logins in the bottom-right corner

### EPJ Swap Alerts
When an EPJ swap occurs, you'll:
- Hear a **double beep** sound (different from login sound)
- See an **orange notification** at the top showing which EPJ they swapped to
- The badge will show both login count and swap count separately

This helps you distinguish between normal activity and potential equipment issues requiring attention.

---

## Tips & Best Practices

1. **Keep the Live Checkouts tab open** during busy periods to monitor equipment usage
2. **Use the Swap feature** instead of force check-in when a driver needs different equipment mid-shift
3. **Add maintenance reasons** to help track recurring issues with specific EPJs
4. **Use bulk user creation** at the start of a season when onboarding multiple drivers
5. **Check the maintenance log** periodically to identify equipment that may need repair or replacement

---

## Troubleshooting

### Page Seems Frozen
Click the **⟳ Refresh** button or refresh your browser.

### "Session Expired" Message
Your login has timed out. The page will automatically refresh in a few seconds—just log back in.

### Can't Remove an EPJ
Make sure the EPJ is not currently checked out. Check the Live Checkouts tab first.

### Driver Can't Log In
Use the User Management page to reset their password or verify their account exists.

---

## Need Help?

Contact your system administrator if you encounter issues not covered in this guide.

---

*Last Updated: January 2026*
