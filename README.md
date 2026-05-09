# Result Analysis API

Service to analyze result `.txt` files and generate an Excel report.

---

## Endpoint

### `POST /analyze_12th`

Upload a `.txt` result file and receive a processed `.xlsx` file.

---

## Request

### Form Data

| Field | Type | Required | Description |
|------|------|------|------|
| `file` | File | Yes | TXT result file |

---

## Example using cURL

```bash
curl -X POST "http://127.0.0.1:8000/analyze_12th" \
-F "file=@result.txt" \
--output result.xlsx