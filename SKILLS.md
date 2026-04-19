# Oracle HC Generator - Specialized Skills & Logic

This document captures the complex logic patterns developed for the Recommendation Engine to ensure consistent behavior and reusability.

## 1. Oracle Patch Date & Age Evaluation (`oracle_version_patch_age`)

### Logic Description
Evaluates the "Database Release Update" status by parsing version numbers and patch application dates from the `CHECK_PATCHES` section of the HTML log.

### Pattern: Robust Date Parsing
- **Format**: `Patch <ID> : applied on [Day] Mon Date Time TZ Year`
- **Regex**: `rf'Patch\s+{ru_patch_id}.*?applied on.*?(?:[a-zA-Z]{3,}\s+)?([a-zA-Z]{3})\s+(\d{1,2}).*?(\d{4})'`
- **Improvement**: Uses a non-capturing group `(?:...)` to optionally skip the day of the week (Wed, Sat, etc.), ensuring the Month is always captured in group 1.

### Pattern: Month-Based Age Calculation
- **Context**: Current system time for logic is **April 2026**.
- **Calculation**: `diff_months = (current_year - patch_year) * 12 + (current_month - patch_month)`.
- **Thresholds**:
    - `> 24 months` -> **Critical**
    - `12 - 24 months` -> **High**
    - `6 - 12 months` -> **Low**
    - `< 6 months` -> **Healthy (No Recommendation)**

---

## 2. Dynamic Report Localization & Styling

### Pattern: Automatic Translation
The engine uses `self.language` to switch between `vi` and `en` keys in the recommendation finding dictionary.
- **Universal Fix**: Replaces "Phụ lục" with "Appendix" in the Final Recommendation table when the language is not Vietnamese.
- **Headers**: Dynamic header row generation for `Recommendations` and `Summary` tables.

### Pattern: Multi-Tiered Severity Styling
Conditional formatting applied to the **Severity** column:
- **Critical / Nghiêm trọng**: `Red` + `Bold`.
- **High / Cao**: `Red` only.
- **Low / Medium**: Black (default).

---

## 3. JSON Configuration Interaction
- **Rule Design**: Use `condition: "oracle_version_patch_age"` to trigger the specialized Python handler.
- **Standardization**: Set `rec_vi`, `rec_en`, etc., to `"Auto-Standardized"` (optional) to use the internal code-based standardized texts for complex branching logic.
