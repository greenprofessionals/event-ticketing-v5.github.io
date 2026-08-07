# Event Ticketing Platform V5

V5 is a clean standalone build. Do not overwrite the working `multi-event-tickets` deployment while testing it.

## Repositories / folders

- `event-ticketing-v5/` — client configuration portal, administration, voucher distribution, claimant ticket rendering.
- `event-gate-v5/` — separate gate-only module for QR check-in, guest lookup, and supervisor exception handling.
- Both modules use the same NEW Google Sheet and the same NEW Apps Script Web App backend.

## System flow

1. System Owner initializes a new Sheet and creates access users.
2. Event Admin creates an Event ID shell and a private client configuration link.
3. Client opens `config.html?event=...&key=...`, fills the embedded Google Form, uploads branding, previews ticket tiers/colors, revises if needed, and approves.
4. Event Admin activates the approved event.
5. Event Admin generates voucher batches by tier and optional suggested chapter/group.
6. Distributor opens one batch link and sends vouchers by Email, WhatsApp, Text, or contacts the recipient by Phone Call.
7. Recipient opens the unique voucher, enters claimant details, and receives the branded QR ticket.
8. Gate Staff uses the separate `event-gate-v5` module to scan or search and check in guests.
9. Gate Supervisors can handle walk-ins and reverse an accidental check-in. Admins handle transfer, QR reissue, revocation, payments, reporting, backups, and audit.

## New Google Sheet / Apps Script setup

1. Create a NEW blank Google Sheet, e.g. `Event Ticketing Platform V5`.
2. Extensions > Apps Script.
3. Paste `Code.gs` from this folder.
4. Run `setupV5System()` once and authorize it.
5. Run `bootstrapOwner()` once and create the System Owner passcode.
6. Run `createClientConfigurationForm()` once. It creates the dedicated Google Form, response sheet, and an installable form-submit trigger.
7. Optional: use `addAccessUser()` from Apps Script, or use Admin > Access Management after deployment.
8. Deploy > New deployment > Web app > Execute as Me > Who has access: Anyone.
9. Copy the `/exec` URL.

## Configure the websites

In both folders, copy `site-config.example.js` to `site-config.js` and paste the SAME new `/exec` endpoint.

Ticketing example:

```js
window.EVENT_TICKETING_CONFIG = {
  ENDPOINT: 'https://script.google.com/macros/s/.../exec',
  GATE_URL: 'https://greenprofessionals.github.io/event-gate-v5/',
  SUPPORT_EMAIL: '',
  SUPPORT_PHONE: ''
};
```

Gate example:

```js
window.EVENT_TICKETING_CONFIG = {
  ENDPOINT: 'https://script.google.com/macros/s/.../exec',
  TICKETING_URL: 'https://greenprofessionals.github.io/event-ticketing-v5/'
};
```

Upload `event-ticketing-v5/` and `event-gate-v5/` as separate GitHub Pages subfolders.

## Important operating rules

- Do not expose the client configuration link publicly. It contains an event-specific configuration token.
- Public ticket claiming is voucher-only.
- An event must be `Client Approved`, then activated by an authorized admin before vouchers can be generated/claimed.
- QR codes remain black on white regardless of tier color.
- Tier color is optional. Blank/invalid colors receive deterministic system defaults.
- Gate Staff never receives event configuration, voucher generation, or finance administration rights.
- System Owner should create a separate access record/passcode for each operator rather than sharing one passcode.
- Use Event Backup before major changes and after each event closes.

## Event lifecycle

`Draft -> Client Submitted -> Preview Ready -> Client Approved -> Active -> Closed -> Archived`

A new Google Form submission after approval resets the event to `Client Submitted` so changes must be previewed and approved again.
