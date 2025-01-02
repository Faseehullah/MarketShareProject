import pandas as pd
import matplotlib.pyplot as plt
import math
from adjustText import adjust_text

def load_data(file_path, sheet_name):
    """Load data from the specified Excel sheet."""
    data = pd.read_excel(file_path, sheet_name=sheet_name)
    print("Columns in the dataset:", data.columns.tolist())
    return data

def standardize_brand(brand):
    """
    Convert brand name to uppercase (no replacements for typos).
    """
    if pd.isnull(brand):
        return None
    return str(brand).strip().upper()

def aggregate_brand_workloads(data, brand_cols, workload_cols):
    """
    Sum up total workloads for each brand across given columns.
    
    :param data: A pandas DataFrame with brand and workload columns.
    :param brand_cols: List of brand column names (e.g., ["IA Brand 1", "IA Brand 2", "IA Brand 3"])
    :param workload_cols: List of workload column names (e.g., ["Workload - Brand 1", "Workload - Brand 2", "Workload - Brand 3"])
    :return: A dictionary {brand_name: total_samples}
    """
    brand_workloads = {}

    # Iterate over each row
    for _, row in data.iterrows():
        # For each brand-workload pair
        for bcol, wcol in zip(brand_cols, workload_cols):
            brand = row.get(bcol, None)
            workload = row.get(wcol, 0)

            # Standardize brand and ensure workload is a valid number
            if pd.notnull(brand) and pd.notnull(workload):
                brand_std = standardize_brand(brand)
                try:
                    workload_val = float(workload)
                except ValueError:
                    workload_val = 0.0

                if workload_val > 0:
                    brand_workloads[brand_std] = brand_workloads.get(brand_std, 0) + workload_val

    return brand_workloads

def calculate_market_share(brand_workloads):
    """
    Given {brand: total_workload}, compute {brand: percentage}.
    """
    total = sum(brand_workloads.values())
    if total == 0:
        return {}
    return {
        brand: (load / total) * 100
        for brand, load in sorted(brand_workloads.items(), key=lambda x: x[1], reverse=True)
    }

def plot_market_share_pie(market_share, title="Market Share", threshold=1.0):
    """Generate a pie chart for market share with callouts."""
    if not market_share:
        print(f"No market share data to plot for {title}.")
        return

    brands = list(market_share.keys())
    shares = list(market_share.values())

    # Combine shares below threshold into 'OTHERS'
    filtered_brands = []
    filtered_shares = []
    other_share = 0.0

    for brand, share in zip(brands, shares):
        if share >= threshold:
            filtered_brands.append(brand)
            filtered_shares.append(share)
        else:
            other_share += share

    if other_share > 0:
        filtered_brands.append("OTHERS")
        filtered_shares.append(other_share)

    plt.figure(figsize=(10, 8))
    wedges, _, _ = plt.pie(
        filtered_shares,
        labels=None,  # We'll add custom labels
        autopct="",
        startangle=140,
        textprops=dict(color="black"),
        wedgeprops=dict(edgecolor="white"),
    )

    # Add labeled callouts outside each wedge
    texts_to_adjust = []
    for wedge, share_val, brand_name in zip(wedges, filtered_shares, filtered_brands):
        angle = (wedge.theta2 - wedge.theta1) / 2 + wedge.theta1
        x = math.cos(math.radians(angle)) * 1.2 * wedge.r
        y = math.sin(math.radians(angle)) * 1.2 * wedge.r
        text = plt.text(
            x,
            y,
            f"{brand_name} ({share_val:.1f}%)",
            ha="center",
            va="center",
            fontsize=9,
            bbox=dict(boxstyle="round,pad=0.3", fc="white", edgecolor="black", alpha=0.6),
        )
        texts_to_adjust.append(text)

    # Dynamically adjust text to avoid overlaps
    adjust_text(texts_to_adjust, arrowprops=dict(arrowstyle="->", color="gray", lw=0.5))

    plt.title(title, fontsize=14)
    plt.legend(
        wedges,
        filtered_brands,
        title="Brands",
        loc="center left",
        bbox_to_anchor=(1, 0, 0.5, 1),
        fontsize=10
    )
    plt.tight_layout()
    plt.show()

def save_to_excel(market_share, output_file, sheet_name="Market Share"):
    """
    Save the market share dictionary to an Excel file.
    """
    if not market_share:
        print("No data to save.")
        return

    df = pd.DataFrame(list(market_share.items()), columns=["Brand", "Market Share (%)"])
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Saved market share data to {output_file} (sheet: {sheet_name}).")

if __name__ == "__main__":
    # 1. Load the data
    file_path = r"E:\FAS NTK\Survey\2024\Visual Code - Market Survey\New Working Jan 2025\Input Jan 2025\market_working_south_only - V1_1.xlsx"
    sheet_name = "IA"  # Update to the correct sheet name
    data = load_data(file_path, sheet_name)

    # 2. Aggregate brand workloads
    brand_cols = ["IA Brand 1", "IA Brand 2", "IA Brand 3"]      # Update as needed
    workload_cols = ["Workload - Brand 1", "Workload - Brand 2", "Workload - Brand 3"]  # Update as needed
    brand_workloads = aggregate_brand_workloads(data, brand_cols, workload_cols)
    print("Brand Workloads:", brand_workloads)

    # 3. Calculate market share
    market_share = calculate_market_share(brand_workloads)
    print("\nMarket Share (%):")
    for b, s in market_share.items():
        print(f"{b}: {s:.2f}%")

    # 4. Plot the pie chart (optional)
    plot_market_share_pie(market_share, title="Immunoassay Market Share", threshold=1.0)

    # 5. Save to Excel
    output_file = r"E:\FAS NTK\Survey\2024\Visual Code - Market Survey\New Working Jan 2025\Output Jan 2025\Market_Share_Output_South_only.xlsx"
    save_to_excel(market_share, output_file, sheet_name="IA Market Share")
