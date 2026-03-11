# Labour Market Stats Briefing Dashboard

**Automated monthly dashboard and narrative report for UK labour market statistics**

---

## Project Structure

```
labour-market-stats-brief/
├── utils/
│   ├── calculations.R    # Main orchestration - sources all sheets
│   ├── config.R          # Reference months and comparison periods
│   ├── helpers.R         # Common helper functions
│   ├── word_output.R     # Word document generation
│   └── README.md         # This file
├── sheets/
│   ├── lfs.R             # Labour Force Survey metrics
│   ├── payroll.R         # Payrolled employees
│   ├── vacancies.R       # Job vacancies
│   ├── wages_nominal.R   # Nominal wage growth
│   ├── wages_cpi.R       # CPI-adjusted wages
│   ├── redundancy.R      # LFS redundancy rates
│   ├── hr1.R             # HR1 redundancy notifications
│   ├── days_lost.R       # Working days lost
│   ├── sector_payroll.R  # Sector-specific payroll
│   ├── inactivity_reasons.R  # Inactivity breakdown
│   ├── summary.R         # Executive summary narrative
│   ├── top_ten_stats.R   # Top 10 statistics narrative
│   └── excel_audit.R     # Excel dashboard generation
└── templates/
    └── dashboard_template.docx  # Word template
```

---

## How to Run

### From RStudio (recommended)

1. Open the project in RStudio
2. Make sure your working directory is the **project root** (not utils/)
3. Open `utils/word_output.R`
4. Click **Source** button

### From R console

```r
setwd("/path/to/labour-market-stats-brief")
source("utils/word_output.R")
```

---

## Configuration

Edit `utils/config.R` to change the reference month:

```r
manual_month <- "dec2023"
manual_month_hr1 <- "nov2023"
```

---

## Path Notes

The code uses relative paths that work when:
- Working directory is project root
- You source `utils/word_output.R`

The `word_output.R` script changes working directory to `utils/` before sourcing `calculations.R`, so:
- `calculations.R` sources files using paths relative to `utils/` (e.g., `../sheets/lfs.R`)
- Template and output paths are relative to project root
