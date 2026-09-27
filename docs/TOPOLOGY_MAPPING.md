# Topology Mapping & Data Dictionary

This document details the relational architecture, object mapping logic, and data dictionary used by `TeamsVoice-GoldenSnapshot`.

---

## 1. Relational Call Flow Architecture

A typical voice application path traverses multiple parent-child entities in Microsoft Teams:

```text
[ PSTN DID / Carrier ]
          │
          ▼
   Resource Account (Application Instance)
          │
          ▼
   Auto Attendant (AA)
     ├── Time Zone / Language / Operator / Authorized Users
     ├── Business Hours Call Flow
     │     └── Menu (DTMF Options: 1..9, 0, *, #)
     │           │
     │           ▼
     │      Call Queue (CQ)
     │        ├── Presence-Based Routing ($true / $false)
     │        ├── Conference Mode / Distribution Method
     │        ├── Agent Roster (Expanded Users + Groups)
     │        ├── Overflow Rules (Threshold -> Action -> Target)
     │        └── Timeout Rules (Duration -> Action -> Target)
     │
     └── After-Hours & Holiday Call Handling Associations