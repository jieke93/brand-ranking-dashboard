import subprocess, os

os.chdir(r'c:\AI실습\VS CODE\UQ')

# Step 1: git add
cmds = [
    ['git', 'add', 'dashboard.py', 'requirements.txt', 'brands_config.json', 'runtime.txt'],
    ['git', 'add', '*.json', '*.py'],
    ['git', 'add', 'image_archive/'],
    ['git', 'add', 'product_images/'],
    ['git', 'add', 'product_images_hd/'],
    ['git', 'add', '.streamlit/config.toml'],
    ['git', 'add', 'auto_deploy.ps1'],
    ['git', 'add', 'update_cloud.bat'],
]

for cmd in cmds:
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
    if r.returncode != 0 and 'did not match' not in r.stderr:
        print(f'WARN: {cmd} -> {r.stderr.strip()}')

print('=== ADD COMPLETE ===')

# Step 2: Show staged files
r2 = subprocess.run(['git', 'diff', '--cached', '--stat'], capture_output=True, text=True, encoding='utf-8')
print(r2.stdout)
print('=== STAGED FILES END ===')

# Step 3: Commit
r3 = subprocess.run(['git', 'commit', '-m', '데이터 업데이트 2026-03-30'], capture_output=True, text=True, encoding='utf-8')
print('=== COMMIT OUTPUT ===')
print(r3.stdout)
if r3.stderr:
    print('COMMIT STDERR:', r3.stderr)
print(f'COMMIT EXIT CODE: {r3.returncode}')

# Step 4: Push
r4 = subprocess.run(['git', 'push', 'origin', 'master'], capture_output=True, text=True, encoding='utf-8')
print('=== PUSH OUTPUT ===')
print(r4.stdout)
if r4.stderr:
    print('PUSH STDERR:', r4.stderr)
print(f'PUSH EXIT CODE: {r4.returncode}')
