# Care AI &mdash; Healthcare Member Assistant & Clinical Operations Cockpit

[![Live Demo](https://img.shields.io/badge/Live_Demo-Care_AI_on_Vercel-5009B5?style=for-the-badge&logo=vercel)](https://carelon-care-ai.vercel.app)
[![License: MIT](https://img.shields.io/badge/License-MIT-01BBBA?style=for-the-badge)](LICENSE)
[![WCAG](https://img.shields.io/badge/Accessibility-WCAG_2.1_AA-43B8F3?style=for-the-badge)](#accessibility)
[![Synthetic Data](https://img.shields.io/badge/Data-100%25_Synthetic-794CFF?style=for-the-badge)](#synthetic-data--governance)

> **Care AI** is an enterprise-grade AI healthcare member assistant and clinical operations platform designed for health insurers and managed care organizations. It combines deterministic rule-based clinical safety filters, document-grounded policy explanations, reading-level adaptation, and zero-repetition clinical escalation handoffs.

---

## 🌟 Live Demo

Experience the live, interactive prototype deployed on Vercel:  
🔗 **[https://carelon-care-ai.vercel.app](https://carelon-care-ai.vercel.app)**

---

## 🏗️ System Architecture & Four Core Modules

Care AI is architected into four cohesive modules connecting member self-service with backend clinical operations:

### 1. Member Experience (`View 1`)
- **Dual Perspective Switcher:** Instant toggle between **Member view** (clean, jargon-free conversational experience with source citations and literacy adaptation) and **Agent view** (operational cockpit with real-time Session Intelligence).
- **5 Presentation Quick Actions:**
  - `Coverage`: Instant diagnostic imaging and benefits lookup.
  - `Claim Status`: Denial reason breakdown with plain-language explanation and 90-day appeal windows.
  - `Prior Approval`: Procedure authorization rules with real-time status tracker (`#PA-88219`).
  - `Deductible`: Live accumulator progress bar and remaining deductible calculation.
  - `Talk to a Person`: Immediate warm transfer to registered nurse or benefits specialist.
- **Reading-Level Adaptation:** Dynamic literacy adaptation toggleable between **4th Grade (Plain English)**, **6th Grade (Standard)**, and **Technical (Detailed)**, with an interactive `[Explain simpler]` action.
- **Care AI Rx Prescription Integration:** Prescription upload interface with interactive mail-order medication template and zero-copay home delivery fulfillment.

### 2. Product & Ops Dashboard (`View 2`)
- **Operational KPIs:** First-Contact Resolution (FCR: 61.3%), Cost per Contact ($1.84 automated), Clinical Recall (99.4%), Escalation Rate (14.2%), and Hallucination Gate (<0.1%).
- **Longitudinal Trends & Deflection:** Month-over-month telemetry tracking inquiries, automated containment, and member CSAT.
- **Equity Disparity Monitor:** Cross-segment parity tracking across Commercial, Medicare, Medicaid, and Spanish-speaking cohorts with automated disparity flags and plain-language template remediation workflows.
- **Human-in-the-Loop Quality Review Queue:** Operational queue enabling clinical supervisors to verify and commit grounded responses or flag exceptions for model fine-tuning.

### 3. Product Roadmap & Backlog (`View 3`)
- **0 to 12-Month Phased Rollout:**
  - `M0–M2`: Phase 1 &mdash; Discovery & Requirements (Gate G1: Scope Approved)
  - `M1–M3`: Phase 2 &mdash; Data Readiness & Design (Gate G2: Design Signed Off)
  - `M2–M6`: Phase 3 &mdash; MVP Build & Integration
  - `M5–M7`: Phase 4 &mdash; Testing & Compliance (Gate G3: Ready for Go-Live)
  - `M7–M9`: Phase 5 &mdash; Pilot & Feedback (Gate G4: Scale Decision)
  - `M9+`: Phase 6 &mdash; Enterprise Scale & Multi-LOB Expansion
- **MoSCoW Prioritization Board:** Clear breakdown across Must-Have, Should-Have, Could-Have, and Won't-Have (Now).
- **Interactive RICE Scoring Table:** Dynamic backlog prioritization model ($$\text{RICE} = \frac{\text{Reach} \times \text{Impact} \times \text{Confidence}}{\text{Effort}}$$) seeded with realistic enterprise epics.

### 4. Agent & Nurse Handoff Console (`View 4`)
- **Zero-Repetition Transfer:** When a member escalates, receiving clinicians and customer care specialists receive a structured handoff dossier containing:
  - Member demographic profile, risk tier, and active plan line.
  - Initial question and exact escalation trigger.
  - Sources checked (Summary of Benefits clauses, real-time accumulator state).
  - Safety screen status and urgency flags.
  - 3-bullet transcript summary eliminating member re-explanation fatigue.
  - Clinician-recommended next steps.
- **Future Extension (Proactive Care Management):** Four-factor targeting engine (Risk × Impactability × Reachability × Timing) with post-discharge 48h fast-lane outreach worklist and demographic fairness monitor.

---

## 🛡️ Deterministic Clinical Safety Architecture

Care AI enforces a strict **two-tier rule-based safety intercept** that executes *before* any language model generation or RAG lookup:

```
[ Member Query Input ]
          |
          v
+-------------------------------------------------------+
|        Tier 1: Emergency Rule-Based Screen            |
|  (chest pain, can't breathe, stroke, severe bleeding, |
|   suicidal distress [988 Lifeline], overdose)         |
+-------------------------------------------------------+
     |                                  |
 [TRIGGERED]                         [PASSED]
     |                                  |
     v                                  v
+-------------------------------+  +-------------------------------------+
| 1. Block generative AI answer |  |     Tier 2: Symptom Safety Screen   |
| 2. Immediate 911 / ER guidance|  | (fever, pain, cough, rash, nausea,  |
| 3. Connect to live nurse line |  |  "I feel sick", dizziness)          |
| 4. Update session telemetry   |  +-------------------------------------+
+-------------------------------+       |                         |
                                   [TRIGGERED]                 [PASSED]
                                        |                         |
                                        v                         v
                       +----------------------------------+  +--------------------------+
                       | 1. Strictly zero medical advice  |  | Standard Cognitive Path  |
                       | 2. Danger signs & 911 screening  |  | - Grounded RAG Benefits  |
                       | 3. Offer 24/7 Nurse Line         |  | - Claims Denial Appeals  |
                       | 4. Offer $0 Telehealth / Urgent  |  | - Prior Authorization    |
                       +----------------------------------+  | - Deductible Progress    |
                                                             +--------------------------+
```

---

## 🧪 Interactive One-Click Demo Scenarios

The prototype includes 5 pre-configured demonstration scenarios accessible via the **⏱ Demo Walkthrough** button:

| Scenario | Persona | Query | Expected System Behavior |
| :--- | :--- | :--- | :--- |
| **1. Coverage & Pre-Auth** | **Alex** (First-Time Member) | *"Is an MRI covered on my plan?"* | Informs diagnostic imaging coverage, 20% coinsurance, prior authorization requirement, and deductible accumulator check. |
| **2. Claims Denial Appeal** | **Maria** (Family Caregiver) | *"Why was my son's ENT visit denied?"* | Explains out-of-network provider denial in plain English, highlights 90-day appeal timeline, and offers agent appeal initiation. |
| **3. Prior Approval** | **Priya** (HDHP + HSA) | *"Do I need approval for my knee surgery?"* | Confirms prior approval requirement, details provider clinical note submission, and opens interactive `#PA-88219` tracker. |
| **4. Emergency Intercept** | **David** (Chronic Condition) | *"I've had chest pain all morning. Is the ER covered?"* | Rule-based safety screen intercepts (0ms), blocks AI response, provides 911/ER guidance, clarifies $0 approval barrier, and bridges nurse line. |
| **5. Symptom Guardrail** | Any Member | *"I'm having fever."* | Prevents medical advice/diagnosis, screens warning signs, offers 24/7 nurse line, and provides $0 telehealth cost guidance. |

---

## 🎨 Design System & Accessibility

- **Care AI Palette:**
  - Primary Purple: `#5009B5`
  - Violet: `#794CFF`
  - Sky Blue: `#43B8F3`
  - Deep Cyan / Teal: `#01BBBA` & `#0D7C8F`
  - High-Contrast Text: `#2B2B2B` (Light) / `#F8FAFC` (Dark)
- **Accessibility:** Designed to meet **WCAG 2.1 AA** standards (minimum 4.5:1 contrast across all interactive controls, readable typography, and aria roles).
- **Typography:** Modern, legible sans-serif hierarchy powered by `Plus Jakarta Sans`.

---

## 🔒 Synthetic Data & Governance

> [!IMPORTANT]
> **All data presented in this repository and demo prototype is 100% synthetic.**
> No protected health information (PHI) or real member records are utilized. Performance metrics, cost models, and clinical scenarios are illustrative models designed for enterprise evaluation.

---

## 💻 Tech Stack & Deployment

- **Frontend:** Responsive Vanilla HTML5, CSS Grid & Flexbox, SVG vector graphics.
- **Client Runtime:** Modern ES6+ JavaScript (zero external framework overhead, instantaneous 60fps interaction).
- **Deployment:** Continuous production deployment via **Vercel Edge Network**.

---

## 📄 License

Distributed under the MIT License. See [LICENSE](LICENSE) for details.
