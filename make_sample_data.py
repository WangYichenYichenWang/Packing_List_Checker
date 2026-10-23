"""
Creates sample input files in data/ so the checker can be tried out.

The data is invented. Each supplier uses different column headers (as in real
life), and a few typical mistakes are built in on purpose so the checker has
something to find.
"""
from pathlib import Path

import pandas as pd

DATA_DIR = Path(__file__).parent / "data"
PL_DIR = DATA_DIR / "packing_lists"


def main():
    PL_DIR.mkdir(parents=True, exist_ok=True)

    # Supplier A: abbreviated headers (CTNS, N.W., G.W.)
    shp_2401 = pd.DataFrame(
        [
            ["SHP-2401", "MAT-H-090", "Hybrid foam mattress 90x200", "940421", 20, 1, 20, 520, 560],
            ["SHP-2401", "MAT-H-140", "Hybrid foam mattress 140x200", "940421", 15, 1, 15, 555, 600],
            # 30 cartons x 8 pcs = 240, but total says 250
            ["SHP-2401", "PIL-MF-01", "Memory foam pillow", "940490", 30, 8, 250, 288, 330],
            # gross weight lower than net weight
            ["SHP-2401", "TOP-160", "Mattress topper 160x200", "940421", 10, 2, 20, 160, 150],
        ],
        columns=["Shipment No", "Item No", "Description", "HS Code", "CTNS",
                 "Qty per CTN", "Total Qty", "N.W. (kg)", "G.W. (kg)"],
    )

    # Supplier B: long-form headers and HS codes written with dots
    shp_2402 = pd.DataFrame(
        [
            ["SHP-2402", "BED-W-140", "Wooden bed frame 140x200", "9403.50", 12, 1, 12, 540, 600],
            # HS code missing
            ["SHP-2402", "DUV-ALL-200", "All-season duvet 200x200", "", 25, 4, 100, 180, 210],
            # HS code differs from the invoice (940421)
            ["SHP-2402", "MAT-H-180", "Hybrid foam mattress 180x200", "940429", 8, 1, 8, 368, 400],
            # HS code has only 5 digits
            ["SHP-2402", "PIL-MF-01", "Memory foam pillow", "94049", 20, 8, 160, 192, 220],
            # free sample packed but not on the commercial invoice
            ["SHP-2402", "CUS-SMP", "Sample cushion (free of charge)", "940490", 1, 2, 2, 1.0, 1.5],
        ],
        columns=["Shipment ID", "Article No", "Goods Description", "Commodity Code", "No. of Cartons",
                 "Pcs per CTN", "Quantity", "Net Weight (kg)", "Gross Weight (kg)"],
    )

    # Supplier C: standard headers
    shp_2403 = pd.DataFrame(
        [
            # invoice says 30 pcs
            ["SHP-2403", "MAT-H-090", "Hybrid foam mattress 90x200", "940421", 25, 1, 25, 650, 700],
            ["SHP-2403", "PIL-MF-01", "Memory foam pillow", "940490", 15, 8, 120, 144, 165],
            # same line entered twice
            ["SHP-2403", "PIL-MF-01", "Memory foam pillow", "940490", 15, 8, 120, 144, 165],
            # TOP-160 is invoiced but missing here
        ],
        columns=["shipment_id", "sku", "description", "hs_code", "cartons",
                 "qty_per_carton", "total_qty", "net_weight_kg", "gross_weight_kg"],
    )

    invoice = pd.DataFrame(
        [
            ["SHP-2401", "MAT-H-090", "Hybrid foam mattress 90x200", "940421", 20, 149.00, "CN"],
            ["SHP-2401", "MAT-H-140", "Hybrid foam mattress 140x200", "940421", 15, 199.00, "CN"],
            ["SHP-2401", "PIL-MF-01", "Memory foam pillow", "940490", 240, 12.50, "CN"],
            ["SHP-2401", "TOP-160", "Mattress topper 160x200", "940421", 20, 45.00, "CN"],
            ["SHP-2402", "BED-W-140", "Wooden bed frame 140x200", "940350", 12, 120.00, "VN"],
            ["SHP-2402", "DUV-ALL-200", "All-season duvet 200x200", "940440", 100, 22.00, "CN"],
            ["SHP-2402", "MAT-H-180", "Hybrid foam mattress 180x200", "940421", 8, 239.00, "TR"],
            ["SHP-2402", "PIL-MF-01", "Memory foam pillow", "940490", 160, 12.50, "CN"],
            ["SHP-2403", "MAT-H-090", "Hybrid foam mattress 90x200", "940421", 30, 149.00, "TR"],
            ["SHP-2403", "PIL-MF-01", "Memory foam pillow", "940490", 120, 12.50, "CN"],
            ["SHP-2403", "TOP-160", "Mattress topper 160x200", "940421", 10, 45.00, "CN"],
        ],
        columns=["Shipment No", "Item No", "Description", "HS Code", "Qty",
                 "Unit Price (EUR)", "Country of Origin"],
    )

    shp_2401.to_excel(PL_DIR / "PL_SHP-2401.xlsx", index=False)
    shp_2402.to_excel(PL_DIR / "PL_SHP-2402.xlsx", index=False)
    shp_2403.to_excel(PL_DIR / "PL_SHP-2403.xlsx", index=False)
    invoice.to_excel(DATA_DIR / "commercial_invoice.xlsx", index=False)
    print(f"Sample files written to {DATA_DIR}")


if __name__ == "__main__":
    main()
