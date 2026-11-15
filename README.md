# Sarvam Market Sizing & Account Map

Built exclusively for **Sarvam AI's PS-II Strategy & Operations internship**. It answers the first
line of that role: *build bottom-up models for the segments we are entering, and find where the next
set of revenue comes from*. It also includes a scored map of target accounts, so sales spends time on
the right ones.

**File:** `sarvam_market_sizing.xlsx` (tabs: Summary, Sources, Inputs, Voice_BFSI,
Dubbing_Content_Studio, Account_Map). 144 live formulas, 0 errors. Rebuild with `python build.py`.

**Facts kept separate from assumptions:** every input is marked **FACT** (green fill, with a source ID
from the Sources tab) or **ASSUMPTION** (yellow fill, blue text, with the reasoning next to it). Change
any assumption and every tab updates.

## Results (annual, base case)
| Segment | Market | Range | Sarvam 3-year obtainable |
| --- | --- | --- | --- |
| **Samvaad voice agents in Indian BFSI** | **₹491 crore (~$56M)** | ₹295–687 crore | ₹25 crore (5% share) |
| **Content Studio dubbing in India** | **₹31 crore (~$3.6M)** | | ₹4.7 crore (15% share) |

**The voice-agent pool is about 16 times the dubbing pool.**

### How the voice-agent number is built
`accounts × contact-centre seats per account × talk minutes per seat × share a voice agent can handle × price`

| Account type | Accounts (FACT) | Share of the market |
| --- | --- | --- |
| NBFC middle layer | 421 | 28% |
| Public sector banks | 12 | 20% |
| Private sector banks | 21 | 14% |
| NBFC upper layer | 15 | 13% |
| Life insurers | 26 | 11% |
| Non-life insurers (incl. health) | 33 | 9% |
| Small finance, regional rural and foreign banks | 84 | 5% |

- **Facts:** account counts from RBI and IRDAI lists; price ₹3.00 a minute (Sarvam's Business plan).
- **Assumptions:** seats per account, 300 days × 300 talk minutes per seat, 25% of minutes handled by a
  voice agent (15% low, 35% high), 20% enterprise discount.

### How the dubbing number is built
`buyers × share that dub × source minutes per buyer × target languages × price`
- **Facts:** 908 private satellite TV channels (Ministry of I&B); 1,900+ films released in 2025
  (FICCI-EY); price ₹38 a minute per language (Growth plan).
- **Assumptions:** OTT, education, creator and enterprise buyer counts, and minutes dubbed.
- Education (35%), creators (23%) and TV channels (21%) are the biggest parts; films are under 3%.

## Takeaways
1. **Voice agents in BFSI are the bigger rupee pool,** because call minutes are huge and every minute
   is billed. NBFCs (41%) and public-sector banks (20%) make up most of it.
2. **Dubbing is smaller at ₹30–38 a minute,** so Content Studio's case rests on fewer, larger accounts
   and on adjacent products (voice cloning, text to speech), not volume alone.
3. **Two inputs move the voice market most:** seats per account and the share of minutes a voice agent
   can handle. Check both in five customer calls before setting targets.
4. **Account map:** the top five are State Bank of India, LIC, Canara Bank, Bajaj Finance and Shriram
   Finance. All combine large call volumes, regional-language customers and a reason to prefer an
   Indian, sovereign AI provider.

## Account map method
Score = pool size 35% + regional-language need 25% + data-sovereignty pull 20% + reachability 20%
(weights editable). The 1–5 scores are my judgment from public information, labelled as inputs to
change, not data. None of the accounts are claimed to be Sarvam customers.

## Limitations
- Seats per account is the weakest input. Company annual reports and contact-centre vendor case
  studies would firm it up.
- The voice model counts contact-centre minutes only, not sales calls from branches or agents.
- Dubbing buyer counts outside TV and films are assumptions.
- Model APIs / Token Factory developer revenue is not sized here; see
  [sarvam-api-funnel-revops-lab](https://github.com/f20220970-debug/sarvam-api-funnel-revops-lab).

## Sources
RBI (banks, NBFC layers), IRDAI via public lists (insurers), Ministry of I&B (TV channels), FICCI-EY
M&E report March 2026 (films, sector size), Sarvam public pricing pages (voice agents, dubbing). Full
links on the Sources tab.

Built by Pavitra Seth (BITS Pilani). Public data only; not affiliated with Sarvam AI.
