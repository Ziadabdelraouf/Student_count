import pandas as pd
import os
import glob
import sys
import matplotlib.pyplot as plt

def generate_schedule_image(schedule, target_code):
    """Generates a PNG image of the student schedule."""
    day_order = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"]
    
    # Filter only days that have sessions
    active_days = [day for day in day_order if schedule[day]]
    
    if not active_days:
        print("No sessions found to plot.")
        return

    # Setup the plot
    fig, ax = plt.subplots(figsize=(10, len(active_days) * 1.5 + 2))
    ax.axis('off')
    
    title = f"Weekly Schedule for Student: {target_code}"
    plt.title(title, fontsize=16, fontweight='bold', pad=20)

    y_pos = 0.9
    for day in day_order:
        sessions = schedule[day]
        if not sessions:
            continue
        
        # Draw Day Header
        ax.text(0.05, y_pos, day.upper(), fontsize=12, fontweight='bold', 
                bbox=dict(facecolor='navy', alpha=0.1, pad=5))
        y_pos -= 0.05
        
        # Draw Sessions
        for session in sorted(sessions):
            ax.text(0.1, y_pos, f"• {session}", fontsize=11, family='monospace')
            y_pos -= 0.04
        
        y_pos -= 0.03 # Extra space between days

    # Save the image
    output_filename = f"schedule_{target_code}.png"
    plt.savefig(output_filename, bbox_inches='tight', dpi=300)
    print(f"\nSchedule image saved as: {output_filename}")

def get_student_schedule(target_code):
    current_dir = os.path.dirname(os.path.abspath(__file__))
    download_dir = os.path.join(current_dir, "downloaded_files_parallel")
    
    excel_files = glob.glob(os.path.join(download_dir, "*.xlsx"))
    
    if not excel_files:
        print("No Excel files found.")
        return

    # Initialize dictionary to group by days
    schedule = {
        "sunday": [], "monday": [], "tuesday": [], "wednesday": [],
        "thursday": [], "friday": [], "saturday": []
    }
    
    # Keep track of valid days for ordering later
    day_order = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"]

    print(f"Scanning {len(excel_files)} files for Student Code: {target_code}...")

    for excel_file in excel_files:
        filename = os.path.basename(excel_file)
        try:
            df = pd.read_excel(excel_file)
            
            if 'Code' not in df.columns:
                continue

            # Check if student exists in this file (exact match)
            students_in_file = set(df['Code'].dropna().astype(str).str.strip())
            
            if str(target_code).strip() in students_in_file:
                
                # 1. Remove extension
                name_no_ext = os.path.splitext(filename)[0]
                
                # 2. Split by underscore
                # Format: {stuff}_{day}_{time}_{uuid}
                parts = name_no_ext.split('_')
                
                # We need at least 3 parts (Day, Time, UUID) for the logic to hold
                if len(parts) >= 3:
                    # The day is the 3rd item from the end
                    # Example: [... 'sunday', '11-1', 'uuid']
                    day_str = parts[3].lower()
                    
                    # The "Clean Name" is everything EXCEPT the last part (the UUID)
                    clean_name = "_".join(parts[:-1])
                    
                    if day_str in schedule:
                        schedule[day_str].append(clean_name)
                    else:
                        # Fallback for unexpected day names
                        # You can create a 'misc' key if needed, or print a warning
                        pass
                        
        except Exception as e:
            print(f"Error processing {filename}: {e}")

    # Print the formatted output
    print("\nSearch Results:\n")
    for day in day_order:
        sessions = schedule[day]
        if sessions:
            # Center the day name with dashes
            header = f"-------{day}-------"
            print(header)
            
            # Sort sessions alphabetically for cleaner look
            for session in sorted(sessions):
                print(session)

    generate_schedule_image(schedule, target_code)

def main():
    if len(sys.argv) > 1:
        code_input = sys.argv[1]
    else:
        code_input = '1230040' 

    get_student_schedule(code_input)

if __name__ == "__main__":
    main()