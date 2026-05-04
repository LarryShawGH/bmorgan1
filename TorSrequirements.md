# ThinkorSwim Automated Trading Indicator — Requirements

## Overview

Develop a custom ThinkorSwim (ThinkScript) study/indicator that plots key price levels and time-based visual elements on the chart to support intraday trading decisions.

---

## 1. Session & Time Configuration

- **Trading Hours:** 9:30 AM – 4:00 PM Eastern Time
- **Overnight Session:** 4:00 PM (prior day close) – 9:30 AM (current day open)
- **Opening Range Window:** 9:30 AM – 10:00 AM (first 30 minutes of RTH)
- All time references must respect the Eastern Time zone

---

## 2. Price Levels to Calculate & Display

### 2.1 Previous Day High/Low
- Calculate the **high** and **low** of the previous regular trading session
- Display as horizontal lines extending across the current session
- **Previous Day High color:** Red
- **Previous Day Low color:** Green

### 2.2 Overnight High/Low
- Calculate the **high** and **low** of the overnight session (4:00 PM prior day → 9:30 AM current day)
- Display as horizontal lines
- **Overnight High color:** Light Green
- **Overnight Low color:** Light Red

### 2.3 Opening Range High/Low
- Calculate the **high** and **low** of the first 30 minutes of the RTH session (9:30 – 10:00 AM)
- Lines should be finalized at 10:00 AM and remain fixed for the rest of the session
- **Opening Range High color:** Distinct Green (differentiated from Previous Day Low green)
- **Opening Range Low color:** Distinct Red (differentiated from Previous Day High red)

---

## 3. Visual Elements

### 3.1 Previous Day High/Low Box
- Draw a shaded rectangle spanning from `previousDayLow` to `previousDayHigh` across the current session
- Include a **center dividing line** at the midpoint of the box
  - Color: White
  - Label: Display the dollar/price value at the midpoint

### 3.2 Half-Hour Vertical Lines
- Draw vertical lines at every 30-minute interval from 9:30 AM to 4:00 PM
  - Intervals: 9:30, 10:00, 10:30, 11:00, 11:30, 12:00, 12:30, 1:00, 1:30, 2:00, 2:30, 3:00, 3:30, 4:00
- Each vertical line must include:
  - **Time label** (e.g., "10:00 AM") displayed on the line
  - **Price/dollar amount** at the time the line is drawn
  - **Color:** White

---

## 4. Color Coding Summary

| Element              | Color               |
|----------------------|---------------------|
| Previous Day High    | Red                 |
| Previous Day Low     | Green               |
| Overnight High       | Light Green         |
| Overnight Low        | Light Red           |
| Opening Range High   | Distinct Green      |
| Opening Range Low    | Distinct Red        |
| Box Center Line      | White               |
| Half-Hour Lines      | White               |

All colors should be visually distinct from one another to avoid ambiguity on the chart.

---

## 5. Alerts

- Trigger an alert when price **crosses above or below** any of the following levels:
  - Previous Day High
  - Previous Day Low
  - Overnight High
  - Overnight Low
  - Opening Range High
  - Opening Range Low
- Alerts should be configurable via ThinkorSwim's native alert system

---

## 6. Settings / User Inputs (Optional but Recommended)

Expose the following as user-configurable inputs in the study settings panel:

- Color preferences for each level (all 6 lines + box)
- Toggle visibility of each level independently (show/hide)
- Toggle box display on/off
- Toggle half-hour lines on/off
- Ability to adjust the opening range window duration (default: 30 minutes)

---

## 7. Performance & Backtesting

- The indicator logic should be compatible with ThinkorSwim's **OnDemand** (paper trading/replay) mode for historical review
- Where possible, structure the code to support **Strategy** mode so trade entries/exits based on level breaks can be backtested using the Strategy Report

---

## 8. Technical Constraints

- Language: **ThinkScript** (native ThinkorSwim scripting language)
- Must function correctly on intraday chart timeframes (1-min, 2-min, 5-min recommended)
- All horizontal lines must be drawn using `plot` or `AddChartBubble`/`DrawLine` as appropriate
- Vertical lines can be implemented using `AddVerticalLine()`
- Price labels can be implemented using `AddChartBubble()` or `AddLabel()`
- Must handle pre-market data gracefully (lines should not distort due to pre-market price action)

---

## 9. Deliverables

1. A single `.ts` ThinkScript file containing the full study
2. Import instructions for loading into ThinkorSwim
3. Brief inline comments in the script explaining each section
