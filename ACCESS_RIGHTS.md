# V5 Access Rights

| Role | Client Configuration | Event/Voucher Admin | Guest & Payment Admin | Gate Check-In | Gate Supervisor | Reports/Audit | Access Management |
|---|---|---|---|---|---|---|---|
| SYSTEM_OWNER | All events | Full | Full | All events | Yes | Full | Full |
| EVENT_ADMIN | Assigned events | Full | Full | Assigned events | Yes | Assigned events | No |
| FINANCE | No configuration changes | No voucher generation | Payment/search on assigned events | No | No | Operational/financial dashboard | No |
| GATE_SUPERVISOR | No | No | Gate-only guest view | Assigned active events | Walk-ins + undo check-in | Gate operational summary | No |
| GATE_STAFF | No | No | Minimal gate guest view | Assigned active events | No | Minimal gate operational summary | No |
| CLIENT | Own private configuration token only | No | No | No | No | No | No |
| DISTRIBUTOR | No | Own batch link only | No | No | No | No | No |
| CLAIMANT | No | Own voucher/ticket only | No | Presents ticket | No | No | No |

Permissions are enforced in the Apps Script backend. Hiding a button in the browser is not treated as authorization.
