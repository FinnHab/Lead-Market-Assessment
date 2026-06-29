# Replication Package — The EU agri-food sector as a lead market for green hydrogen

Finn Haberkost, Sibusisiwe Khuzwayo, Adela Marian, Rainer Quitzow

Research Institute for Sustainability (RIFS) at GFZ, Potsdam, Germany.

Package version: 1.0 (2026-06-29).

This package contains the scripts and result spreadsheets underlying the four-dimensional lead-market assessment (scale, affordability, domestic market autonomy, administrative feasibility) presented in the manuscript. The scripts are the authoritative source of all numerical values; the spreadsheets in `outputs/` are snapshots produced by these scripts.

---

## Repository structure

```
Replication_Package/
├── README.md
├── inputs/                                 (12 files, see Inputs table below)
├── scripts/
│   ├── _styling.py                         shared visual styling module
│   ├── 01_scale_results.py
│   ├── 02_scale_index.py
│   ├── 03_affordability_results.py
│   ├── 04_affordability_index.py
│   ├── 05_price_transmission.py
│   ├── 06_ipr_export_intensity.py
│   ├── 07_dma_index.py
│   ├── 08_admin_feasibility_results.py
│   └── 09_admin_feasibility_index.py
└── outputs/
    ├── 01_scale_results.xlsx
    ├── 02_scale_index.xlsx
    ├── 03_affordability_results.xlsx
    ├── 04_affordability_index.xlsx
    ├── 05_price_transmission.xlsx
    ├── 06a_ipr_export_intensity_2023.xlsx  reference year used in manuscript
    ├── 06b_ipr_export_intensity_2024.xlsx  robustness, follow-up year
    ├── 07_dma_index.xlsx
    ├── 08_admin_feasibility_results.xlsx
    └── 09_admin_feasibility_index.xlsx
```

---

## File-to-figure mapping

| Script | Output | Manuscript reference |
|---|---|---|
| 01_scale_results.py | 01_scale_results.xlsx | Figure 1 (Sankey + regulatory reach); Methods, Scale |
| 02_scale_index.py | 02_scale_index.xlsx | Figure 6 (radar, scale axis); Methods, Scale |
| 03_affordability_results.py | 03_affordability_results.xlsx | Figure 2 (green premium cascade); Methods, Affordability |
| 04_affordability_index.py | 04_affordability_index.xlsx | Figure 6 (radar, affordability axis); Methods, Affordability |
| 05_price_transmission.py | 05_price_transmission.xlsx | Figure 4 (price transmission); Methods, DMA |
| 06_ipr_export_intensity.py | 06a_ipr_export_intensity_2023.xlsx, 06b_ipr_export_intensity_2024.xlsx | Figure 3 (trade intensity bubble); Methods, DMA |
| 07_dma_index.py | 07_dma_index.xlsx | Figure 6 (radar, DMA axis); Methods, DMA |
| 08_admin_feasibility_results.py | 08_admin_feasibility_results.xlsx | Figure 5 (market concentration); Methods, Administrative feasibility |
| 09_admin_feasibility_index.py | 09_admin_feasibility_index.xlsx | Figure 6 (radar, feasibility axis); Methods, Administrative feasibility |

---

## Pipeline structure

The pipeline consists of four independent dimension stacks (Scale, Affordability, DMA, Administrative feasibility). Within each stack the upstream "results" script reads raw inputs and produces a sector-level results workbook; the downstream "index" script consumes those results as a small set of hard-coded sector arrays and produces the normalised index.

```
Scale stack:        01_scale_results.py        ->  01_scale_results.xlsx
                                                       |
                                                       v
                    02_scale_index.py          ->  02_scale_index.xlsx

Affordability stack: 03_affordability_results.py  ->  03_affordability_results.xlsx
                                                       |
                                                       v
                    04_affordability_index.py  ->  04_affordability_index.xlsx

DMA stack:           05_price_transmission.py   ->  05_price_transmission.xlsx
                    06_ipr_export_intensity.py  ->  06a_/06b_*.xlsx
                                                       |
                                                       v
                    07_dma_index.py            ->  07_dma_index.xlsx

Admin feasibility:   08_admin_feasibility_results.py  ->  08_admin_feasibility_results.xlsx
                                                       |
                                                       v
                    09_admin_feasibility_index.py  ->  09_admin_feasibility_index.xlsx
```

**Self-contained downstream scripts.** Scripts 02, 04, 07 and 09 are self-contained: they hold the per-sector inputs as hard-coded constants synchronised manually from the corresponding upstream results workbook. They do not read the upstream `.xlsx` at run-time. If you change an upstream parameter and re-run 01/03/05/06/08, you must re-derive the downstream constants by hand for the change to propagate to the index workbooks. Scripts 03 and 08 are similarly self-contained at the input level: their numerical inputs are hard-coded from external sources cited in the script header and in the `Sources & Methods` sheet of the corresponding output workbook where present.

---

## Inputs

All files in `inputs/` are redistributable under their respective open licences. Eurostat data are covered by the Commission Reuse Decision (2011/833/EU); the EU-27 nitrogen-application file is the Dryad-published dataset (CC-By); DG AGRI Crop Market Observatory data are public DG AGRI publications; FAOSTAT is CC-By; UN Comtrade is public.

| File | Used by | Type | Description | Primary source |
|---|---|---|---|---|
| eurostat_apro_crop_areas_2023.xlsx | 01 | raw | EU-27 crop areas, production and yields by crop, 2023 | Eurostat apro_cpsh1 |
| ifa_nitrogen_application_by_crop.xlsx | 01 | raw | Nitrogen application rates and planted area by country and crop | Dryad CC-By |
| dg_agri_crop_balance_sheets.xlsx | 01 | raw | DG AGRI commodity balance sheets, EU-27, 2023/24 | DG AGRI Crop Market Observatory |
| comext_imports_eu27_2023.xlsx | 01 | pre-aggregated | Extra-EU import volumes by product, sector-aggregated, 2023 | Eurostat COMEXT |
| faostat_exporter_yields.xlsx | 01 | pre-aggregated | Crop yields for major EU import-origin countries, 2021-2023 average | FAOSTAT QCL |
| price_transmission_panel.xlsx | 05 | pre-aggregated | Monthly EU vs world unit-value panel, Jan 2010 - Dec 2023 | UN Comtrade + Eurostat APRO/APRI |
| cn4_nace_template.xlsx | 06 | concordance | CN4 to NACE sector concordance with per-CN4 basis flags | Eurostat correspondence tables |
| PRODCOM_CN_correspondance_table_2024_EN.xlsx | manual verification | reference | Official PRODCOM to CN8 correspondence, 2024 | Eurostat RAMON |
| apri_eu_prices_components.csv | 06 | raw | Eurostat APRI producer-price index, EU-27 components | Eurostat apri_pi23_outm |
| estat_apro_cpsh1.tsv.gz | 06 | raw | Eurostat APRO crop production bulk, gzipped | Eurostat apro_cpsh1 bulk |
| prodcom_overlap_resolution_tables_2024.xlsx | 06 | concordance | PRODCOM overlap resolution for codes mapped to multiple CN4 | DG GROW PRODCOM |
| sbs_market_concentration_raw.xlsx | manual verification | reference | Eurostat SBS market-concentration raw extraction | Eurostat SBS |

`eurostat_apro_crop_areas_2023.xlsx` and `estat_apro_cpsh1.tsv.gz` are the same underlying dataset (Eurostat APRO) at two granularities: the xlsx is the pre-aggregated EU-27 summary used by 01, the gzipped TSV is the country-level bulk dump used by 06.

The hard-coded constants in 01_scale_results.py (Scale chain 1, animal-feed allocation) come from published sources cited in-script: FEFAC 2024 yearbook (compound feed volumes), IFA/IFDC Global Fertilizer Use by Crop 2017/18, FAO LEAP / GLEAM 2017 (on-farm vs ICF shares), and IEAGHG 2023 / EC DG ECFIN / CBAM Regulation 2023/956 (referenced from 03).

### External raw data not included in this package

Two raw Eurostat dumps are required to re-run script 06 end-to-end but are too large to bundle with the replication package. They are not redistributed here. Reviewers can download them directly from Eurostat at no cost; the dataset codes the script expects are listed below.

| File expected by script 06 | Dataset code | Eurostat download page |
|---|---|---|
| `full_2023*.dat` and `full_2024*.dat` (~4 GB per year, 12 monthly files each) | COMEXT extra-EU-27 monthly trade | https://ec.europa.eu/eurostat/databrowser/bulk?lang=en&selectedTab=fileComext&breadcrumbFilter=COMEXT_DATA%2FPRODUCTS&searchFilter=full_partxixu_v2_2023 |
| `estat_ds-059358.tsv` (~290 MB) | PRODCOM annual production statistics (`DS-059358`) | https://ec.europa.eu/eurostat/databrowser/bulk?lang=en&searchFilter=ds-059358 |

To re-run 06 end-to-end, download these two datasets from Eurostat, place the monthly COMEXT `full_<YYYYMM>.dat` files and the PRODCOM TSV directly under `inputs/`, then run 06 from the `scripts/` directory. Without these files, 06 will not run; the result workbooks `06a_ipr_export_intensity_2023.xlsx` and `06b_ipr_export_intensity_2024.xlsx` in `outputs/` still allow full inspection of the computed values.

---

## Contact

Finn Haberkost — finn.haberkost@rifs-potsdam.de

Research Institute for Sustainability (RIFS) at GFZ, Potsdam.
