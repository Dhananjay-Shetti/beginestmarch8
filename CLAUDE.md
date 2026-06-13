# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Beginest** is a meeting room booking portal — a static HTML/vanilla JS frontend backed by Google Apps Script (GAS) deployed as a web app, with Google Sheets as the data store. There is no build step, no package manager, and no local dev server.

## Development

### Local Development
Open any `.html` file directly in a browser. All backend calls go to the deployed GAS macro via the `scriptUrl` constant defined near the top of each file. Update `scriptUrl` in each HTML file to point to your own deployed GAS deployment if testing against a different backend.

### Backend Deployment
`Code.gs` must be deployed via the Google Apps Script editor (`script.google.com`). Changes to `Code.gs` require a new GAS deployment version to take effect in production.

There are no build, lint, or test commands — this project has none.

## Architecture

### File Roles
| File | Purpose |
|------|---------|
| `index.html` | Shell/host page. Manages auth state (24h session in localStorage), renders the header, and swaps iframe content between pages. |
| `register.html` | Login and first-time registration. Password validation enforced client-side before the GAS call. |
| `MeetingRoomLandingPage.html` | Main booking UI. Two modes: Preference (date/time → available rooms) and Single Room (direct room view/book). |
| `MyBookings.html` | Upcoming and past 30-day booking history; cancel bookings with credit refund. |
| `Hamburger.html` | Sidebar/nav component — persistent 360px panel on desktop, overlay drawer on mobile. |
| `Code.gs` | All backend logic: auth, room queries, booking, credit deduction, availability checking, concurrency locking. |

### Communication Patterns
- **Frontend → Backend:** `fetch()` against the GAS `scriptUrl` with `?action=<endpoint>` query params (GET) or POST body.
- **Iframe coordination:** `index.html` and `Hamburger.html` communicate via `window.postMessage` — e.g., refreshing the credit display after a booking.
- **Session:** Auth state stored in `localStorage` (`whatsappNumber`, `companyName`, `userName`, `loginTime`).

### Data Model (Google Sheets)
- **Data sheet:** User records — WhatsApp number is the primary key; password stored as SHA-256 hash.
- **HarborMeetingRoomMap sheet:** Room inventory — each row has a JSON blob with capacity, amenities, image URL, floor, and availability flag.
- **BookingLog sheet:** Each booking row: booking ID, user info, room, date, time slot string (e.g. `"15:00 to 17:00"`), harbor.
- **CompanyCreditsMap sheet:** Per-company remaining and total credits by harbor.

### Booking Credit System
Credits deducted = `room_capacity × duration_hours`. Cancellations refund the same amount. Credits are company-scoped, not per-user.

### Conflict Prevention
Double-booking is prevented in two layers:
1. **Client-side:** `checkEndCandidates` GAS endpoint checks multiple end-time candidates in one request; UI grays out unavailable slots.
2. **Server-side:** Before writing a booking, `Code.gs` acquires `LockService.getScriptLock()` (waits up to 5 s) and re-validates availability.

### Slot Representation
Time slots are stored as strings like `"15:00 to 17:00"`. Client code expands these into 30-minute half-hour segments for overlap detection. The date picker blocks weekends and past dates client-side.

### Caching
GAS uses `CacheService` with a ~30-second TTL on booking data to reduce Sheets read load. Room list data is cached in `localStorage`.
