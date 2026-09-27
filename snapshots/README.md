# TeamsVoice-GoldenSnapshot

[![PowerShell](https://img.shields.io/badge/PowerShell-7.2+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Platform](https://img.shields.io/badge/Platform-Microsoft%20Teams%20Phone-0078D4.svg)](https://learn.microsoft.com/microsoftteams/teams-phone)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

An enterprise-grade configuration baseline and disaster-recovery auditing engine for **Microsoft Teams Phone System**. 

`TeamsVoice-GoldenSnapshot` traverses your tenant's Auto Attendants (AA) and linked Call Queues (CQ) to produce an immutable, row-per-route golden snapshot. It solves the critical enterprise operational challenge of **untracked configuration drift**, allowing Teams voice administrators to compare point-in-time routing baselines during service degradation, misrouting tickets, or audit cycles.

---

## 🎯 What Problem Does It Solve?

In enterprise Microsoft Teams telephony, routing logic naturally fractures over time:
* An administrator modifies a keypad (DTMF) menu option without logging the change.
* An offboarded user is disabled in Entra ID, creating an orphaned call queue.
* Timeout durations, conference modes, or presence-based routing settings are adjusted during troubleshooting and never reverted.
* A department submits an escalation: *"Key 2 on our main line worked last month, but now it fails."*

Because native administration consoles lack point-in-time configuration diffing, diagnosing these issues often requires reconstructed guesswork. **TeamsVoice-GoldenSnapshot** provides a standardized, queryable snapshot capturing resource accounts, phone numbers, DTMF routing keys, presence settings, exception thresholds, and expanded agent rosters in a single normalized CSV.

---

## 🚀 Key Features

* **Row-per-Route Normalization:** Flattens multi-tiered Auto Attendant call trees into structured, single-line routing records without column explosion.
* **Full Topology Correlation:** Maps Auto Attendants to downstream Call Queues, Resource Accounts, and phone numbers (DIDs).
* **Deep Agent Expansion:** Resolves both direct user assignments and nested Entra ID Security Groups into readable member rosters.
* **Directory Status Auditing:** Enriches agent rosters with real-time Entra ID account status (`[Active]` vs `[Disabled]`).
* **Anti-Throttling Architecture:** Implements in-memory hashtable batch caching to protect against Microsoft Graph HTTP 429 rate limits.
* **Disaster Recovery Ready:** Can be archived monthly or integrated into version-controlled baseline repositories.

---

## 📂 Repository Layout

```text
TeamsVoice-GoldenSnapshot/
├── .gitignore
├── LICENSE
├── README.md
├── docs/
│   └── TOPOLOGY_MAPPING.md          # Technical documentation and schema dictionary
├── snapshots/                       # Local directory for monthly golden CSVs (git-ignored)
│   └── README.md                    # Storage and disaster recovery operational guide
└── src/
    └── Export-TeamsVoiceGoldenSnapshot.ps1  # Primary data capture engine