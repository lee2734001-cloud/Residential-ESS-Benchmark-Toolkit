# Residential ESS Benchmark Toolkit

A reusable benchmarking repository for residential energy storage systems.

## Current module

**Template 01 – Product Dimension Comparison**

Benchmark condition:

- 10 kW three-phase inverter
- 1 battery pack
- Official product images only
- No AI redraw
- Dimensions use total system envelope: Width × Depth × Height

Current brands:

- Huawei LUTERRA S2
- Sigenergy SigenStor Neo
- FoxESS HQ3 + EQ5000
- SolarEdge Nexis three-phase + Nexis Battery
- EcoFlow PowerOcean 2 Three-Phase

## Repository structure

```text
Assets/      Official product images
Database/    Excel/CSV benchmark data
Master/      PowerPoint master/template
Scripts/     Automatic PPT generator
Output/      Generated benchmark presentations
Docs/        Data schema and workflow guides
.github/     GitHub Actions cloud automation
```

## Generate the PowerPoint

```bash
pip install -r requirements.txt
python Scripts/benchmark_generator.py
```

The output is written to:

```text
Output/Residential_ESS_Dimension_Benchmark.pptx
```

## Run without installing a GitHub client

1. Open this repository in a browser.
2. Upload the extracted repository files with **Add file → Upload files**.
3. Open **Settings → Actions → General**.
4. Under **Workflow permissions**, select **Read and write permissions**.
5. Open the **Actions** tab and run **Generate benchmark PPT**.

The workflow generates the PPT in GitHub's cloud environment and commits the updated file to `Output/`.

## Update a product

1. Replace the image in `Assets/<Brand>/`.
2. Update `Database/products.xlsx` or `Database/products.csv`.
3. Upload the changed files through the GitHub browser.
4. The GitHub Action regenerates the output automatically.

## Data conventions

- `Width_mm`: X axis
- `Depth_mm`: Y axis
- `Height_mm`: Z axis
- `Footprint_m2 = Width × Depth / 1,000,000`
- `Volume_m3 = Width × Depth × Height / 1,000,000,000`

## Version

`v1.0.0` – initial repository package.
