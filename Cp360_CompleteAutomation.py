import subprocess

def run_script(script_name):
    try:
        print(f"Running {script_name}...")
        result = subprocess.run(["python", script_name], check=True)
        print(f"{script_name} completed successfully.\n")
    except subprocess.CalledProcessError as e:
        print(f"Error running {script_name}: {e}")
        exit(1)  # Exit the program if there's an error

# Run the first script
run_script("LumberJack_v10.py")

# Run the second script
run_script("CP360Automation_v4.py")