import pandas as pd
import os
from dateutil.parser import parse
from datetime import datetime
import calendar
import zipfile
import shutil

class TimeValidator:
    '''Class for timesheet validation operations'''

    def __init__(self):
        self.results = []

    def standardize_column_names(self, df):
        '''Standardize column names across different sheets.'''
        df = df.copy()
        # Standardize Hours column
        for col in df.columns:
            if "Duration (in hrs)" in col or "Hours" in col:
                df.rename(columns={col: "Hours"}, inplace=True)

        if "Description" in df.columns:
            df.rename(columns={"Description": "Sheet Name"}, inplace=True)

        return df

    def validate(self, df):
        '''Validate timesheet data according to business rules.'''
        
        df = self.standardize_column_names(df)
        
        required_columns = ["Client", "Sheet Name", "Hours"]
        for col in required_columns:
            if col not in df.columns:
                df[col] = None
        
        df["Status"] = "Valid"
        df["Flag"] = ""
        
        for index, row in df.iterrows():
            client = str(row["Client"]).strip() if pd.notna(row["Client"]) else ""
            sheet_name = str(row["Sheet Name"]).strip() if pd.notna(row["Sheet Name"]) else ""
            hours = row["Hours"]

            # --- Weekend flag logic ---
            date_val = row.get("Date", "")
            weekday = ""
            try:
                if pd.notna(date_val):
                    weekday = pd.to_datetime(date_val).strftime("%A")
            except Exception:
                pass
            if weekday in ["Saturday", "Sunday"]:
                # If weekend and hours is not empty/zero, raise flag
                if pd.notna(hours) and hours not in [0, "0", "", None]:
                    df.at[index, "Flag"] = (df.at[index, "Flag"] + "⚠ Weekend filled; "
                                            if df.at[index, "Flag"] else "⚠ Weekend filled; ")

            is_leave_type = any(leave_type.lower() in client.lower() for leave_type in ["leave", "holiday", "weekend"])

            if is_leave_type:
                # For leave types, hours should be 0 or empty
                if pd.notna(hours) and hours != 0:
                    df.at[index, "Status"] = "Leave/Holiday should be 0 or empty"
                else:
                    # Leave with 0 hours or empty is valid
                    df.at[index, "Status"] = "Valid"
            else:
                if pd.notna(hours):
                    if isinstance(hours, (int, float)):
                        if hours == 4:
                            df.at[index, "Status"] = "Half-day detected"
                            df.at[index, "Flag"] = "⚠ Half-Day Alert"
                        elif hours != 8:
                            df.at[index, "Status"] = f"Full working day should be 8 hrs, found {hours} hrs"
                        # If hours is 8, it's already marked as Valid
                    else:
                        df.at[index, "Status"] = "Invalid Hours Format"
                else:
                    df.at[index, "Status"] = "Missing hours for a working day"

        for index, row in df.iterrows():
            description = str(row["Sheet Name"]).strip() if pd.notna(row["Sheet Name"]) else ""  # Assuming "Sheet Name" is the description column

            if not description:  
                df.at[index, "Flag"] = df.at[index, "Flag"] + "⚠ Blank Description; " if df.at[index, "Flag"] else  "⚠ Blank Description; "

        for index, row in df.iterrows():
            date_str = str(row["Date"])  

            try:
                
                parsed_date = parse(date_str)
                
                df.at[index, "Date"] = parsed_date.strftime("%Y-%m-%d")  

            except ValueError:
                
                df.at[index, "Date"] = None
        
        result_df = df[["Client", "Date", "Sheet Name", "Hours", "Status", "Flag"]]

        return result_df

    def create_summary(self, validated_sheets):
        '''Create a summary sheet with serial numbers.'''
        
        summary_columns = ["S.No", "File Name", "Sheet Name", "Hours", "Review"]
        summary_data = []
        
        s_no = 1
        
        for sheet_name, df in validated_sheets.items():
            
            total_hours = df["Hours"].sum() if "Hours" in df.columns and df["Hours"].notna().any() else 0
            
            issues = []
            
            if any(df["Status"] == "Half-day detected"):
                issues.append("Contains half-days")
            
            non_standard_entries = df[df["Status"].str.contains("Full working day should be 8 hrs", na=False)]
            if not non_standard_entries.empty:
                issues.append("Has non-standard hours")
            
            if any(df["Status"] == "Leave/Holiday should be 0 or empty"):
                issues.append("Incorrectly logged leave/holiday")

            # Check for missing hours
            if any(df["Status"] == "Missing hours for a working day"):
                issues.append("Missing hours entries")

            # Check for invalid formats
            if any(df["Status"] == "Invalid Hours Format"):
                issues.append("Invalid hour format")

            if any(df["Flag"].str.contains("Blank Description", na=False)):  # Using str.contains for flexibility
                issues.append("Has blank descriptions")

            if any(df["Flag"].str.contains("Weekend filled", na=False)):
                issues.append("Contains weekend entries")    

            # Combine issues into review message
            review_message = ", ".join(issues) if issues else "OK"

            # Add to summary data
            summary_data.append({
                "S.No": s_no,
                "File Name": sheet_name,  # Using sheet name as file name for now
                "Sheet Name": sheet_name,
                "Hours": total_hours,
                "Review": review_message
            })

            # Increment serial number
            s_no += 1

        # Create summary DataFrame
        summary_df = pd.DataFrame(summary_data, columns=summary_columns)

        total_hours_row = pd.DataFrame([{'S.No': '', 'File Name': 'Total', 'Sheet Name': '', 'Hours': summary_df['Hours'].sum(), 'Review': ''}])
        summary_df = pd.concat([summary_df, total_hours_row], ignore_index=True)

        return summary_df

    def run(self, file_path):
        '''Run validation on all sheets in an Excel file.'''
        print(f"Validating file: {file_path}")
        try:
            # Load all sheets
            sheets_dict = pd.read_excel(file_path, sheet_name=None)

            # Dictionary to store validated data
            validated_sheets = {}

            # Process each sheet
            for sheet_name, df in sheets_dict.items():
                print(f"Processing sheet: {sheet_name}")
                validated_sheets[sheet_name] = self.validate(df)

            # Extract file name from path
            file_name = os.path.basename(file_path)

            # Create summary
            summary = self.create_summary(validated_sheets)

            # Update File Name in summary to use actual file name
            summary["File Name"] = file_name

            # Store results
            result = {
                "file_path": file_path,
                "validated_sheets": validated_sheets,
                "summary": summary,
                "success": True  # Add success flag
            }

            self.results.append(result)
            return result

        except Exception as e:
            print(f"Error processing {file_path}: {str(e)}")
            return {"success": False, "error": str(e)}
        
class OutputManager:
    """Class for handling output operations"""

    def __init__(self, output_dir, archive_dir, base_validation_dir):
        self.output_dir = output_dir
        self.archive_dir = archive_dir
        self.base_validation_dir = base_validation_dir
        self.current_validation_dir = None
        
        # Create directories if they don't exist
        for directory in [output_dir, archive_dir, base_validation_dir]:
            if not os.path.exists(directory):
                os.makedirs(directory)

    def set_validation_directory(self, validation_number=None):
        """Set the current validation directory to use.

        Args:
            validation_number: Integer or string to specify which validation folder to use
                               If None, uses the base validation directory
        """
        if validation_number is None:
            self.current_validation_dir = self.base_validation_dir
        else:
            validation_folder = f"validation{validation_number}"
            self.current_validation_dir = os.path.join(self.base_validation_dir, validation_folder)

        # Create the directory if it doesn't exist
        if not os.path.exists(self.current_validation_dir):
            os.makedirs(self.current_validation_dir)

        return self.current_validation_dir

    def get_summary_file_path(self):
        """Get the path to the summary file based on current validation directory"""
        if not self.current_validation_dir:
            self.set_validation_directory()

        # Use a different summary filename based on the validation folder
        if self.current_validation_dir == self.base_validation_dir:
            return os.path.join(self.current_validation_dir, "validation_summary.xlsx")
        else:
            folder_name = os.path.basename(self.current_validation_dir)
            return os.path.join(self.current_validation_dir, f"{folder_name}_summary.xlsx")

    def save_validated_data(self, validation_result, validation_number=None, output_path=None):
        """Save validated data to a new Excel file in the specified validation directory.

        Args:
            validation_result: Dictionary containing validation results
            validation_number: Which validation folder to use (None for base folder)
            output_path: Custom output path (optional)
        """
        # Ensure validation result has success key
        if not validation_result.get("success", True):  # Default to True if key doesn't exist
            print("Validation result contains errors, cannot save.")
            return None

        # Set validation directory based on validation_number
        self.set_validation_directory(validation_number)

        file_name = os.path.basename(validation_result["file_path"])

        # Use a timestamp in filename to avoid overwrites of same file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        subfolder_name = f"validation_{timestamp}"
        subfolder_path = os.path.join(self.current_validation_dir, subfolder_name)
        os.makedirs(subfolder_path, exist_ok=True)
        
        #base_filename = os.path.splitext(file_name)[0] 
        output_path = os.path.join(subfolder_path, f"{file_name}")

        summary_df = validation_result["summary"]  # Get the summary DataFrame
        summary_path = os.path.join(subfolder_path, "validation_summary.xlsx")  # Path for summary file
        summary_df.to_excel(summary_path, index=False)  # Save the summary

        if output_path is None:
            # Create an appropriate filename based on the validation directory
            if self.current_validation_dir == self.base_validation_dir:
                output_name = f"validation_{timestamp}_{file_name}"
            else:
                folder_name = os.path.basename(self.current_validation_dir)
                output_name = f"{folder_name}_{timestamp}_{file_name}"

            output_path = os.path.join(self.current_validation_dir, output_name)

        print(f"Saving validated data to {output_path}...")
        try:
            # Get data from validation result
            validated_sheets = validation_result["validated_sheets"]
            summary_df = validation_result["summary"]

            with pd.ExcelWriter(output_path) as writer:
                # Then write each validated sheet
                for sheet, df in validated_sheets.items():
                    df.to_excel(writer, sheet_name=sheet, index=False)

            print(f"Validation complete. File saved as {output_path}")
            print(f"A summary sheet with serial numbers has been added")

            # Add file to summary tracking for both the specific validation folder and master summary
            self.add_to_summary_tracking(validation_result, validation_number)

            # Also add to the master summary in the base directory
            if validation_number is not None:
                self.add_to_master_summary(validation_result)

            return output_path
        except Exception as e:
            print(f"Error saving file: {str(e)}")
            return None

    def add_to_summary_tracking(self, validation_result, validation_number=None):
        """Add validation results to a summary tracking file in the specific validation directory."""
        # Get the path to the appropriate summary file based on validation number
        self.set_validation_directory(validation_number)
        summary_path = self.get_summary_file_path()

        file_name = os.path.basename(validation_result["file_path"])
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Create or load summary tracking file
        if os.path.exists(summary_path):
            try:
                summary_tracking = pd.read_excel(summary_path)
            except:
                # Create new if reading fails
                summary_tracking = pd.DataFrame(columns=["S.No", "File Name", "Sheet Name", "Hours", "Review", "Validation Date"])
        else:
            summary_tracking = pd.DataFrame(columns=["S.No", "File Name", "Sheet Name", "Hours", "Review", "Validation Date"])

        # Create new entries for each sheet in the validation result
        new_entries = []

        # Get the last S.No value and increment from there
        start_sno = 1
        if not summary_tracking.empty and "S.No" in summary_tracking.columns:
            try:
                start_sno = summary_tracking["S.No"].max() + 1
            except:
                start_sno = 1

        # Add entries for each validated sheet
        sno = start_sno
        for sheet_name, df in validation_result["validated_sheets"].items():
            # Calculate total hours for this sheet
            total_hours = 0
            if "Hours" in df.columns:
                total_hours = df["Hours"].sum()

            # Determine review message
            review_msg = ""
            if any(df["Status"] != "Valid"):
                issues = []
                if any(df["Status"] == "Half-day detected"):
                    issues.append("Contains half-days")
                if any(df["Status"].str.contains("Full working day should be 8 hrs", na=False)):
                    issues.append("Has non-standard hours")
                if any(df["Status"] == "Leave/Holiday should be 0 or empty"):
                    issues.append("Incorrectly logged leave/holiday")
                if any(df["Status"] == "Missing hours for a working day"):
                    issues.append("Missing hours entries")
                if any(df["Status"] == "Invalid Hours Format"):
                    issues.append("Invalid hour format")
                review_msg = ", ".join(issues)
            else:
                review_msg = "OK"

            new_entries.append({
                "S.No": sno,
                "File Name": file_name,
                "Sheet Name": sheet_name,
                "Hours": total_hours,
                "Review": review_msg,
                "Validation Date": timestamp
            })
            sno += 1

        # Create DataFrame from new entries
        new_entry_df = pd.DataFrame(new_entries)

        # Append to existing summary tracking
        summary_tracking = pd.concat([summary_tracking, new_entry_df], ignore_index=True)

        # Save updated summary tracking
        summary_tracking.to_excel(summary_path, index=False)
        print(f"Summary tracking updated at {summary_path}")

    def add_to_master_summary(self, validation_result):
        """Add validation results to the master summary tracking file in the base validation directory."""
        # Always use the base validation directory for the master summary
        master_summary_path = os.path.join(self.base_validation_dir, "master_validation_summary.xlsx")

        file_name = os.path.basename(validation_result["file_path"])
        validation_folder = os.path.basename(self.current_validation_dir)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Create or load master summary tracking file
        if os.path.exists(master_summary_path):
            try:
                master_summary = pd.read_excel(master_summary_path)
            except:
                # Create new if reading fails
                master_summary = pd.DataFrame(columns=["S.No", "File Name", "Sheet Name", "Hours", "Review",
                                                     "Validation Folder", "Validation Date"])
        else:
            master_summary = pd.DataFrame(columns=["S.No", "File Name", "Sheet Name", "Hours", "Review",
                                                 "Validation Folder", "Validation Date"])

        # Create new entries for each sheet in the validation result
        new_entries = []

        # Get the last S.No value and increment from there
        start_sno = 1
        if not master_summary.empty and "S.No" in master_summary.columns:
            try:
                start_sno = master_summary["S.No"].max() + 1
            except:
                start_sno = 1

        # Add entries for each validated sheet
        sno = start_sno
        for sheet_name, df in validation_result["validated_sheets"].items():
            # Calculate total hours for this sheet
            total_hours = 0
            if "Hours" in df.columns:
                total_hours = df["Hours"].sum()

            # Determine review message
            review_msg = ""
            if any(df["Status"] != "Valid"):
                issues = []
                if any(df["Status"] == "Half-day detected"):
                    issues.append("Contains half-days")
                if any(df["Status"].str.contains("Full working day should be 8 hrs", na=False)):
                    issues.append("Has non-standard hours")
                if any(df["Status"] == "Leave/Holiday should be 0 or empty"):
                    issues.append("Incorrectly logged leave/holiday")
                if any(df["Status"] == "Missing hours for a working day"):
                    issues.append("Missing hours entries")
                if any(df["Status"] == "Invalid Hours Format"):
                    issues.append("Invalid hour format")
                review_msg = ", ".join(issues)
            else:
                review_msg = "OK"

            new_entries.append({
                "S.No": sno,
                "File Name": file_name,
                "Sheet Name": sheet_name,
                "Hours": total_hours,
                "Review": review_msg,
                "Validation Folder": validation_folder,
                "Validation Date": timestamp
            })
            sno += 1

        # Create DataFrame from new entries
        new_entry_df = pd.DataFrame(new_entries)

        # Append to existing master summary tracking
        master_summary = pd.concat([master_summary, new_entry_df], ignore_index=True)

        # Save updated master summary tracking
        master_summary.to_excel(master_summary_path, index=False)
        print(f"Master summary tracking updated at {master_summary_path}")

    def create_zip_archive(self, file_path, zip_path=None):
      """Create a ZIP archive containing the validated file and its summary."""
      if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return None

      summary_path = os.path.join(os.path.dirname(file_path), "validation_summary.xlsx")
      if not os.path.exists(summary_path):
        print(f"Warning: Summary file not found at {summary_path}. Only adding main file.")

      if zip_path is None:
        file_name = os.path.basename(file_path)
        zip_name = f"{os.path.splitext(file_name)[0]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = os.path.join(self.output_dir, zip_name)

      try:
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            # Use input file name + _validated.xlsx for the validated file inside the zip
            base_name, ext = os.path.splitext(os.path.basename(file_path))
            validated_name = f"{base_name}_validated{ext}"
            zipf.write(file_path, validated_name, zipfile.ZIP_DEFLATED)
            # Add the summary file if it exists, always as Validation_Summary.xlsx
            if os.path.exists(summary_path):
                zipf.write(summary_path, "Validation_Summary.xlsx", zipfile.ZIP_DEFLATED)


        print(f"ZIP archive created at {zip_path} (includes summary if found)")
        return zip_path
      except Exception as e:
        print(f"Error creating ZIP archive: {str(e)}")
        return None

    def create_new_version(self, file_path):
        """Create a new version of the sheet and archive the old one."""
        if not os.path.exists(file_path):
            print(f"Error: File not found at {file_path}")
            return None, None

        # Get file info
        file_name = os.path.basename(file_path)
        file_base = os.path.splitext(file_name)[0]
        file_ext = os.path.splitext(file_name)[1]

        # Create archive name with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        archive_name = f"{file_base}_v{timestamp}{file_ext}"
        archive_path = os.path.join(self.archive_dir, archive_name)

        # Copy current file to archive
        try:
            shutil.copy2(file_path, archive_path)
            print(f"Previous version archived as {archive_path}")

            # Return the path to the new version (same as original path)
            return file_path, archive_path
        except Exception as e:
            print(f"Error creating new version: {str(e)}")
            return None, None

    def generate_monthly_template(self, month=None, year=None):
        """Generate a new template for a specific month."""
        # If month or year not provided, use next month
        current_date = datetime.now()
        if month is None:
            # If current month is December, go to January of next year
            if current_date.month == 12:
                month = 1
                year = current_date.year + 1
            else:
                month = current_date.month + 1
                year = current_date.year
        elif year is None:
            year = current_date.year

        # Get month name
        month_name = calendar.month_name[month]

        # Create a new dataframe for the month
        # Get number of days in the month
        _, num_days = calendar.monthrange(year, month)

        # Create data for all days in the month
        data = []
        for day in range(1, num_days + 1):
            date = datetime(year, month, day)
            weekday = date.strftime("%A")

            # Set default values based on weekday
            if weekday in ["Saturday", "Sunday"]:
                client = "Weekend"
                hours = 0
            else:
                client = ""  
                hours = 8    

            data.append({
                "Date": date,
                "Day": weekday,
                "Client": client,
                "Sheet Name": "",  
                "Hours": hours
            })

        # Create dataframe
        monthly_df = pd.DataFrame(data)

        # Generate output file path
        output_name = f"Timesheet_{month_name}_{year}.xlsx"
        output_path = os.path.join(self.output_dir, output_name)

        # Save to Excel
        try:
            monthly_df.to_excel(output_path, index=False, sheet_name=f"{month_name} {year}")
            print(f"Monthly template for {month_name} {year} created at {output_path}")
            return output_path
        except Exception as e:
            print(f"Error creating monthly template: {str(e)}")
            return None

# Example usage
def main():
    # Define local directories
    base_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
    media_dir = os.path.join(base_dir, 'media')
    output_dir = os.path.join(media_dir, 'timesheet_outputs')
    archive_dir = os.path.join(media_dir, 'timesheet_archives')
    validation_dir = os.path.join(media_dir, 'timesheet_validations')
    
    # Create directories if they don't exist
    for directory in [output_dir, archive_dir, validation_dir]:
        if not os.path.exists(directory):
            os.makedirs(directory)
    
    # Initialize classes
    validator = TimeValidator()
    output_manager = OutputManager(output_dir, archive_dir, validation_dir)
    
    print("Timesheet validation tool initialized with local directories.")
    print(f"Output directory: {output_dir}")
    print(f"Archive directory: {archive_dir}")
    print(f"Validation directory: {validation_dir}")
    
    return validator, output_manager

if __name__ == "__main__":
    main()