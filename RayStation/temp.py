import os, re
base = r'T:\Physics\Elekta\Elekta Profiles Comparison\Profile Comparisons'
r = r'^(ELEKTA|SBRT 6MV) vs (E\d Water Tank)'
for f in os.listdir(base):
    m = re.search(r, f)
    if m is not None:
        end = f[m.end():].strip().strip('- SBRT ')
        new_f = m.group(2) + ' vs ' + m.group(1) + ' - ' + end
        f = os.path.join(base, f)
        new_f = os.path.join(base, new_f)
        os.rename(f, new_f)