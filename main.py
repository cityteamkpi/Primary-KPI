# --- Standard Library Imports ---
from datetime import datetime
from zoneinfo import ZoneInfo
import argparse
import sys

# --- Third-Party Imports ---
import functions_framework

# ==========================================
# Input and Output Folder Names
# ==========================================

#Folders
input_folder = "Apricot Report Incoming"
output_folder = "Quarterly KPI Processed Data"

# File configuration is managed within specific processors



# ==========================================
# Google Cloud Function Handler
# ==========================================

@functions_framework.http
def run_my_script(request):
    """
    HTTP Cloud Function entry point.
    Expects a JSON payload: 
    {
        "start_date": "YYYY-MM-DD", 
        "end_date": "YYYY-MM-DD" 
      }
    """
    # 1. Extract task: Check JSON body first, then fallback to URL query parameters
    request_json = request.get_json(silent=True) or {}
    task = request_json.get("task") or request.args.get("task", "all")

    print(f"Executing request - Task: {task}, Payload: {request_json}, QueryArgs: {dict(request.args)}")

    # Default to current month if dates aren't provided
    # Using America/Los_Angeles to ensure PST/PDT is handled correctly
    today = datetime.now(ZoneInfo("America/Los_Angeles"))
    default_start = today.replace(day=1).strftime('%Y-%m-%d')
    default_end = today.strftime('%Y-%m-%d')
    
    start_date = request_json.get("start_date", default_start)
    end_date = request_json.get("end_date", default_end)

    print(f"Starting task '{task}'")

    try:
        results = []

        if task in ["renew", "all"]:
            from renew_programs_processor import run_renew_processing
            print("Running Renew Programs Processor...")
            run_renew_processing(
                input_folder_name=input_folder,
                output_folder_name=output_folder
            )
            results.append("Renew Programs Processing Complete.")

        if task in ["alum", "all"]:
            from alum_processor import run_alum_processing
            print("Running Alumni Processor...")
            msg = run_alum_processing(
                input_folder_name=input_folder,
                output_folder_name=output_folder
            )
            results.append(msg)

        if task in ["alum-gd", "all"]:
            from alum_processor import run_alum_processing
            print("Running Alumni Processor...")
            msg = run_alum_processing(
                start_date=start_date,
                end_date=end_date,
                input_folder_name=input_folder,
                output_folder_name=output_folder
            )
            results.append(msg)

        if task in ["util", "all"]:
            from program_utilization_processor import run_utilization_processing
            print("Running Program Utilization Processor...")
            run_utilization_processing(
                input_folder_name=input_folder,
                output_folder_name=output_folder
            )
            results.append("Program Utilization Processing Complete.")

        if task in ["impact", "all"]:
            from kpi_impact_processor import run_kpi_impact_processing
            print("Running KPI Impact Processor...")
            run_kpi_impact_processing(
                input_folder_name=input_folder,
                output_folder_name=output_folder
            )
            results.append("KPI Impact Processing Complete.")

        if task in ["tp", "all"]:
            from turning_point_processor import run_turning_point_processing
            print("Running Turning Point Processor...")
            run_turning_point_processing(
                input_folder_name=input_folder,
                output_folder_name=output_folder
            )
            results.append("Turning Point Processing Complete.")

        if task in ["atwork", "all"]:
            from atwork_processor import run_atwork_processing
            print("Running @Work Processor...")
            run_atwork_processing(
                input_folder_name=input_folder,
                output_folder_name=output_folder
            )
            results.append("@Work Processing Complete.")

        final_msg = " | ".join(results)
        print(f"Success: {final_msg}")
        return final_msg, 200
    except ValueError as e:
        print(f"Configuration Error: {str(e)}")
        return f"Bad Configuration: {str(e)}", 400
    except FileNotFoundError as e:
        print(f"Resource Error: {str(e)}")
        return f"File Not Found: {str(e)}", 404
    except Exception as e:
        print(f"Error during processing: {str(e)}")
        return f"Processing Error: {str(e)}", 500

# ==========================================
# Local Command-Line Interface (CLI)
# ==========================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Local runner for KPI processors")
    parser.add_argument("--start_date", help="Start date (YYYY-MM-DD)")
    parser.add_argument("--end_date", help="End date (YYYY-MM-DD)")
    parser.add_argument("--task", choices=["renew", "alum", "util", "impact", "tp", "atwork", "all"], default="all", help="Task to run")

    args = parser.parse_args()

    # Default dates if not provided (Local CLI execution)
    today = datetime.now(ZoneInfo("America/Los_Angeles"))
    cli_start = args.start_date or today.replace(day=1).strftime('%Y-%m-%d')
    cli_end = args.end_date or today.strftime('%Y-%m-%d')

    try:
        if args.task in ["renew", "all"]:
            from renew_programs_processor import run_renew_processing
            print("Running Renew Programs Processor...")
            run_renew_processing(
                input_folder_name=input_folder,
                output_folder_name=output_folder
            )

        if args.task in ["alum"]:
            from alum_processor import run_alum_processing
            print(f"Running Alumni Processor ...")
            run_alum_processing(
                input_folder_name=input_folder,
                output_folder_name=output_folder
            )

     #   if args.task in ["alum-gd"]:
     #       from alum_processor import run_alum_processing
     #       print(f"Running Alumni Processor for {cli_start} to {cli_end}...")
     #       run_alum_processing(
     #           start_date=cli_start,
     #           end_date=cli_end,
     #           input_folder_name=input_folder,
     #           output_folder_name=output_folder
     #       )

        if args.task in ["util", "all"]:
            from program_utilization_processor import run_utilization_processing
            print("Running Program Utilization Processor...")
            run_utilization_processing(
                input_folder_name=input_folder,
                output_folder_name=output_folder
            )

        if args.task in ["impact", "all"]:
            from kpi_impact_processor import run_kpi_impact_processing
            print("Running KPI Impact Processor...")
            run_kpi_impact_processing(
                input_folder_name=input_folder,
                output_folder_name=output_folder
            )

        if args.task in ["tp", "all"]:
            from turning_point_processor import run_turning_point_processing
            print("Running Turning Point Processor...")
            run_turning_point_processing(
                input_folder_name=input_folder,
                output_folder_name=output_folder
            )

        if args.task in ["atwork", "all"]:
            from atwork_processor import run_atwork_processing
            print("Running @Work Processor...")
            run_atwork_processing(
                input_folder_name=input_folder,
                output_folder_name=output_folder
            )
    except Exception as e:
        print(f"Error during CLI execution: {e}")
