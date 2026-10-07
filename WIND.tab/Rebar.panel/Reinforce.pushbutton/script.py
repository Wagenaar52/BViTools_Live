import os
import sys
import time
from pathlib import Path
from pyrevit import forms

# Get the rebar.panel directory
panel_dir = Path(__file__).parent.parent


# Get the file path once from user
def getFilePath():
    FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)
    if not FPath:
        raise SystemExit
    return FPath


file_path = getFilePath()

if not file_path:
    print("No file path provided. Exiting.")
    sys.exit()

# Dynamically discover all pushbutton scripts in the panel directory (recursive)
scripts_to_run = []
for pushbutton_dir in panel_dir.rglob('*.pushbutton'):
    if pushbutton_dir.is_dir() and pushbutton_dir.name not in ['Links_99d.pushbutton', 'Reinforce.pushbutton']:
        script_path = pushbutton_dir / 'script.py'
        if script_path.exists():
            scripts_to_run.append({'name': pushbutton_dir.name, 'path': script_path})
scripts_to_run = sorted(scripts_to_run, key=lambda x: x['name'])

# Results tracking
results = {
    'successful': [],
    'failed': [],
    'timings': []
}

# Run each script
for script_info in scripts_to_run:
    script_name = script_info['name']
    script_path = script_info['path']
    start_time = time.time()
    
    try:
        # Execute the script with the file path
        with open(str(script_path), 'r') as f:
            script_code = f.read()
        
        # Create a namespace with file_path variable available
        namespace = {'__file__': str(script_path), 'file_path': file_path, 'forms': forms}

        # Force child scripts to reuse the already selected file path.
        original_pick_file = forms.pick_file
        forms.pick_file = lambda *args, **kwargs: file_path
        try:
            exec(script_code, namespace)
        finally:
            forms.pick_file = original_pick_file

        elapsed = time.time() - start_time
        results['timings'].append({'script': script_name, 'seconds': elapsed})
        
        results['successful'].append(script_name)
        print("[OK] {} executed successfully in {:.2f}s".format(script_name, elapsed))
        
    except Exception as e:
        elapsed = time.time() - start_time
        results['timings'].append({'script': script_name, 'seconds': elapsed})
        results['failed'].append({
            'script': script_name,
            'error': str(e)
        })
        print("[FAIL] {} failed after {:.2f}s: {}".format(script_name, elapsed, str(e)))

# Print final report
print("\n" + "="*60)
print("EXECUTION REPORT")
print("="*60)
print("\nFile Path Used: {}\n".format(file_path))

print("Successful Scripts ({}):".format(len(results['successful'])))
for script in results['successful']:
    print("  [OK] {}".format(script))

print("\nFailed Scripts ({}):".format(len(results['failed'])))
for item in results['failed']:
    print("  [FAIL] {}".format(item['script']))
    print("    Error: {}".format(item['error']))

if results['timings']:
    print("\nSlowest Scripts:")
    for timing in sorted(results['timings'], key=lambda x: x['seconds'], reverse=True)[:5]:
        print("  {}: {:.2f}s".format(timing['script'], timing['seconds']))

print("\n" + "="*60)