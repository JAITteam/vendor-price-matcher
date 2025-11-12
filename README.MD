# Vendor Price Matcher

Automatically match vendor prices from VPL files to OITM files for SAP B1.

## What It Does

- Matches prices from vendor files to your item master data
- Creates DTW-ready files for SAP B1 import
- Handles size mapping (XS → XSM) for specific vendors
- Generates reports showing matched and unmatched items

## Quick Start

**1. Install Python packages:**
```bash
pip install pandas openpyxl
```

**2. Prepare files:**
- OITM file: `V105_OITM.xlsx`
- VPL file: `V105_VPL.xlsx`
- Put both in `C:\Users\it2\Downloads\`

**3. Run:**
```bash
python vendor_price_matcher.py
```

**4. Get results in:**
```
C:\Users\it2\Downloads\results\
```

## File Format

### OITM File
```
Item No.
2278-BLACK-XS
2278-BLACK-S
```

### VPL File
```
Vendor Style | Color | Size | Variable | Price
2278         | BLACK | XS   |          | 12.50
2278         | BLACK | S    |          | 12.50
```

## Configuration

To add vendors that use "G" sizes (XS → XSM):
```python
STYLES_USING_G_SIZES = [
    '2278',    # Your style number
]
```

## Output

- `V105_OITM_Updated.xlsx` - Ready for DTW import
- `Processing_Summary_*.xlsx` - Match rates and removed items

## Troubleshooting

**No files found?**
- Check file names: `V105_OITM.xlsx` and `V105_VPL.xlsx`
- Vendor codes must match

**Low match rate?**
- Check color names match exactly
- Verify Item No. format: `STYLE-COLOR-SIZE`

**Need help?**
- Email: it@jauniforms.com

---

**Version:** 1.0  
**Author:** JA Uniforms IT Department