# Vendor Price Matcher & Discontinued Items Finder

Two automated tools for SAP Business One Item Master Data management.

![Python Version](https://img.shields.io/badge/python-3.7+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

---

## 🛠️ Tools Included

### 1. Vendor Price Matcher (`vendor_price_matcher.py`)

Automatically match vendor prices from VPL files to OITM files for SAP B1.

**What it does:**
- Matches prices from vendor VPL files to OITM (Item Master Data)
- Creates DTW-ready import files with updated U_VendorCost
- Handles size mapping for specific vendors (XS → XSM)
- Generates match reports and identifies unmatched items

**Use when:**
- Updating vendor costs monthly/quarterly
- Processing new vendor price lists
- Bulk updating U_VendorCost field in SAP B1

---

### 2. Discontinued Items Finder (`find_discontinued_items.py`)

Identifies discontinued items and creates deactivation files for SAP B1.

**What it does:**
- Scans vendor DTW/VPL files for items marked as "DISCONTINUED"
- Cross-references with your current OITM (active items)
- Creates DTW-ready deactivation files
- Sets frozenFor=Y and validFor=N for bulk deactivation

**Use when:**
- Vendor discontinues products
- Need to deactivate multiple items at once
- Cleaning up obsolete inventory items

---

## 🚀 Quick Start

### Installation
```bash
# Clone repository
git clone https://github.com/JAITteam/vendor-price-matcher.git
cd vendor-price-matcher

# Install dependencies
pip install -r requirements.txt
```

---

## 📋 Tool 1: Vendor Price Matcher

### Usage

**1. Prepare files:**
- OITM file: `V105_OITM.xlsx`
- VPL file: `V105_VPL.xlsx`
- Place in: `C:\Users\it2\Downloads\`

**2. Run:**
```bash
python vendor_price_matcher.py
```

**3. Results:**
```
C:\Users\it2\Downloads\results\
├── V105_OITM_Updated.xlsx          ← Import via DTW
└── Processing_Summary_*.xlsx       ← Review match rates
```

### File Format

**OITM File:**
```
Item No.
2278-BLACK-XS
2278-BLACK-S
```

**VPL File:**
```
Vendor Style | Color | Size | Variable | Price
2278         | BLACK | XS   |          | 12.50
2278         | BLACK | S    |          | 12.50
```

### Output

- `V###_OITM_Updated.xlsx` - DTW-ready with updated prices
- `Processing_Summary_*.xlsx` - Match rates and removed items

---

## 📋 Tool 2: Discontinued Items Finder

### Usage

**1. Prepare files:**
- OITM file: `V105_OITM.xlsx` (current active items)
- DTW/VPL file: `V105_DTW.xlsx` (with "DISCONTINUED" in Style Name)
- Place in: `C:\Users\it2\Downloads\`

**2. Run:**
```bash
python find_discontinued_items.py
```

**3. Results:**
```
C:\Users\it2\Downloads\results\
├── V105_DEACTIVATE_DTW.xlsx        ← Import via DTW
└── Deactivation_Summary_*.xlsx     ← Review deactivations
```

### File Requirements

**OITM File:**
```
ItemCode
2278-BLACK-XS
2278-BLACK-S
```

**DTW File (must have "Style Name" column):**
```
Vendor Style | Style Name                    | Color | Size
2278         | DISCONTINUED-Style 2278      | BLACK | XS
2278         | DISCONTINUED-Style 2278      | BLACK | S
```

### Output

**Deactivation File Format:**
```
ItemCode         | frozenFor | validFor
2278-BLACK-XS    | Y         | N
2278-BLACK-S     | Y         | N
```

**Import via DTW:**
1. Open DTW (Data Transfer Workbench)
2. Select: Item Master Data template
3. Choose: **UPDATE** mode (not Add)
4. Map columns: ItemCode, frozenFor, validFor
5. Import file

---

## ⚙️ Configuration

### Size Mapping (Vendor Price Matcher)

Edit `vendor_price_matcher.py` to add vendors that use "G" suffix sizes:
```python
STYLES_USING_G_SIZES = [
    '2278',              # Fabian Group
    '3483',              # Rothco
    ('2795', 'SILVER'),  # Vantage Apparel (specific color)
]
```

Size conversions:
```
XS  → XSM
S   → SM
M   → MD
L   → LG
XL  → XLG
2XL → 2XLG
```

---

## 🔧 Troubleshooting

### Vendor Price Matcher

**No files found?**
- Check file naming: `V105_OITM.xlsx` and `V105_VPL.xlsx`
- Vendor codes must match exactly

**Low match rate (<85%)?**
- Verify color names match exactly (BLACK vs BLK)
- Check Item No. format: `STYLE-COLOR-SIZE`
- Ensure size codes are consistent

---

### Discontinued Items Finder

**No discontinued items found?**
- Check DTW file has "Style Name" column
- Verify "DISCONTINUED" text appears in Style Name
- Text search is case-insensitive

**Items not matching?**
- Ensure ItemCode format matches in both files
- Check for typos in style/color/size
- Verify items exist in OITM before deactivation

---

## 📖 Documentation

- Full SOP (Standard Operating Procedure) available on request
- Both tools support batch processing of multiple vendors
- Files automatically process all vendor pairs in input directory

---

## 🤝 Author

**JAITTEAM** - JA Uniforms IT Department  
**Contact:** it@jauniforms.com

---

## 📝 License

This project is licensed under the MIT License.

---

## 🎯 Workflow Summary

### Monthly Price Update Process:
1. Receive vendor price lists
2. Run **Vendor Price Matcher** → Get updated prices
3. Run **Discontinued Items Finder** → Identify discontinued items
4. Review both outputs
5. Import via DTW to SAP B1

### When Vendor Discontinues Products:
1. Vendor sends updated file with "DISCONTINUED" items
2. Run **Discontinued Items Finder**
3. Review deactivation list
4. Import to SAP B1 (sets frozenFor=Y, validFor=N)

---

**Version:** 1.0  
**Last Updated:** November 12, 2025