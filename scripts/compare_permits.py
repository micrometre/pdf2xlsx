import pandas as pd
import argparse

def compare_permits(daily_scans_file, residents_permits_file):
    # Read the CSV files
    daily_scans = pd.read_csv(daily_scans_file)
    residents_permits = pd.read_csv(residents_permits_file)
    
    # Get the plate numbers from both files
    scanned_plates = set(daily_scans['Plate'].dropna().astype(str))
    permit_plates = set(residents_permits['Plate'].dropna().astype(str))
    
    # Find permits that are absent (in residents_permits but not in daily_scans)
    absent_permits = permit_plates - scanned_plates
    
    # Create a DataFrame with the absent permits in original order from residents_permits
    absent_plates_list = [plate for plate in residents_permits['Plate'].dropna().astype(str) if plate in absent_permits]
    absent_df = pd.DataFrame({'Plate': absent_plates_list})
    
    print(f"Total permits in {residents_permits_file}: {len(permit_plates)}")
    print(f"Total plates scanned in {daily_scans_file}: {len(scanned_plates)}")
    print(f"Absent permits: {len(absent_permits)}")
    
    return absent_df

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Compare daily scans with residents permits to find absent permits')
    parser.add_argument('daily_scans', help='Path to daily scans CSV file')
    parser.add_argument('residents_permits', help='Path to residents permits CSV file')
    parser.add_argument('-o', '--output', default='absent_permits.csv', help='Output file name (default: absent_permits.csv)')
    
    args = parser.parse_args()
    
    absent_permits = compare_permits(args.daily_scans, args.residents_permits)
    
    # Save to output file
    absent_permits.to_csv(args.output, index=False)
    print(f"Results saved to {args.output}")
    
    print("\nFirst 10 absent permits:")
    print(absent_permits.head(10) if not absent_permits.empty else "No absent permits found.")
