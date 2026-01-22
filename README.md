## Features

- Automatically merge Shopee product, basic info, and media data
- Convert to Myship batch upload format (single/dual variant products)
- Automatic batch processing (default 100 products per batch)
- Automatically convert image Hash ID to full URL
- Preserve Excel VBA macro support
- File size validation (2MB limit)

## Installation

```bash
# Install dependencies
pip install -r requirements.txt
```

## Configuration

1. Copy the configuration template:
```bash
cp config.example.json config.json
```

2. Edit `config.json` and fill in your store information:
```json
{
  "store": {
    "name": "Your Store Name",
    "description": "Your store description",
    "temperature": "Room Temperature"
  },
  "files": {
    "sales": "sales.xlsx",
    "basicinfo": "basicinfo.xlsx",
    "media": "media.xlsx",
    "template": "myship_batch_01_test.xlsm"
  },
  "batch": {
    "size": 100,
    "max_file_size_mb": 2
  }
}
```

## Usage

### 1. Prepare Shopee Export Data

Export the following files from Shopee Seller Center:
- `sales.xlsx` - Product price, inventory, SKU, specifications
- `basicinfo.xlsx` - Product name, description
- `media.xlsx` - Product images

### 2. Run Conversion

**Method A: Convert products only (without store information)**
```bash
python shopee_to_myship_github.py
```
Output: `myship_upload_batch1.xlsm`, `myship_upload_batch2.xlsm`, ...

**Method B: Generate complete upload file with store information**
```bash
python create_final_upload_github.py
```
Output: `myship_upload_ready.xlsm`

### 3. Upload to Myship

1. Log in to 7-11 Myship backend
2. Select "Batch Upload" feature
3. Upload the generated `.xlsm` file

## File Description

### Core Tools
- `shopee_to_myship_github.py` - Main conversion script (batch processing)
- `create_final_upload_github.py` - Complete upload file generation (with store)
- `config.example.json` - Configuration file template
- `requirements.txt` - Python dependencies

### Configuration
- `config.json` - Your store configuration (do not upload to GitHub)

### Myship Template
- `*.xlsm` - Official Myship batch upload template

## Conversion Rules

- Maximum 100 products per batch (adjustable in `config.json`)
- File size limit: 2MB
- Image conversion: Hash ID to full URL
- Uses "Single Product Import" format
- Preserves all required worksheet structure and VBA macros

## Data Mapping

| Shopee Field | Myship Field |
|-------------|-------------|
| et_title_product_name | Product Name |
| ps_item_cover_image | Product Image (URL) |
| et_title_product_description | Product Description |
| et_title_variation_name | Variant |
| et_title_variation_stock | Quantity |
| et_title_variation_price | Price |

## Important Notes

- `config.json` contains your store information - do not share or upload to GitHub
- Original Excel files contain product data - store them securely
- Check file size before upload (must not exceed 2MB)
- Recommend testing with a small number of products first

## License

MIT License

## Contributing

Pull requests and issues are welcome!

## Privacy & Disclaimer

- This project is provided for personal use only; do not use it for commercial purposes without proper authorization.
- All data you process with these scripts remains your responsibility—handle store credentials and product data securely.
- The authors provide no warranties or guarantees; use at your own risk.
