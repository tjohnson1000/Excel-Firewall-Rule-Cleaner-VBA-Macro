# Excel Firewall Rule Cleaner – VBA Macro

## Overview
This **VBA macro project** automates the cleanup of firewall rule data in Excel, specifically targeting **destination IPs/FQDNs** and **TCP/UDP service fields**. It applies **regular expressions** to parse and normalize firewall rule entries—ideal for InfoSec analysts handling messy data exports.

## Features
- Cleans and validates destination entries:
  - Extracts valid **IPv4 addresses**
  - Detects **FQDNs** ending in `.com` or `.net`
- Normalizes service entries:
  - Filters only valid `tcp-xxxx` or `udp-xxxx` formats
  - Converts service protocols to lowercase
- Cleans entire worksheet in one click
- Built-in message alert after cleaning

## Use Case
Security engineers often export firewall rules into Excel from platforms like Palo Alto, Cisco, or AlgoSec. This macro automates the process of cleaning raw exports for:
- Accurate rule reviews
- Faster analysis
- Reduced human error during parsing

## Core Functions
### 🔍 `CleanDestinationField(rawText As String) → String`
Parses mixed destination data and returns a clean, comma-separated string of valid **IPs and FQDNs**.

### 🔍 `CleanServiceField(rawText As String) → String`
Parses mixed service/port data and returns cleaned `tcp-xxxx` or `udp-xxxx` format.

### 🔁 `Sub CleanFirewallColumns()`
Loops through Excel rows and applies the above functions to each `Destination` and `Service` column entry.

## Sample Regex Patterns Used
```vb
ipRegex.Pattern = "\b(?:\d{1,3}\.){3}\d{1,3}\b"
fqdnRegex.Pattern = "\b[a-zA-Z0-9-]+\.(com|net)\b"
svcRegex.Pattern = "^(tcp|udp)-\d+$"
```

## Repository Structure
```
├── CleanDest_Svc_Macro_Full.bas    # VBA module file
├── README.md
```

## Conclusion
This Excel macro streamlines firewall log and rule cleanup, saving analysts hours of manual work. It's a practical automation script for cybersecurity teams and SOC analysts working with bulk data in Excel.

---
**Author:** Travis Johnson  
**Company:** 10Digit Solutions LLC  
**GitHub Repository:** [Add link when available]
