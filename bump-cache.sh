#!/usr/bin/env bash
# ---------------------------------------------------------------------------
# Ανανεώνει το ?v=... σε κάθε τοπικό αρχείο (εικόνες, ήχος) μέσα στο index.html.
# Το ?v= είναι το «αποτύπωμα» (hash) του αρχείου: όταν αλλάξεις μια φωτογραφία,
# αλλάζει και το URL, οπότε οι browsers κατεβάζουν υποχρεωτικά τη νέα έκδοση
# αντί να δείχνουν την παλιά από την cache τους.
#
# Χρήση: κάθε φορά που αλλάζεις φωτογραφία ή μουσική, τρέξε:
#     ./bump-cache.sh
# και μετά κάνε commit + push.
# ---------------------------------------------------------------------------
set -euo pipefail
cd "$(dirname "$0")"

python3 - <<'PY'
import hashlib, pathlib, re, sys

html = pathlib.Path("index.html")
src = html.read_text(encoding="utf-8")

EXT = r"jpg|jpeg|png|webp|gif|svg|avif|ico|mp3|mp4|m4a|webm|woff2?"
pattern = re.compile(
    r"""(?<=["'(])([A-Za-z0-9._/-]+\.(?:%s))(?:\?v=[A-Za-z0-9]+)?(?=["')])""" % EXT
)

seen, missing = {}, []

def repl(m):
    name = m.group(1)
    path = pathlib.Path(name)
    if not path.is_file():
        missing.append(name)
        return m.group(0)
    digest = hashlib.md5(path.read_bytes()).hexdigest()[:8]
    seen[name] = digest
    return f"{name}?v={digest}"

out = pattern.sub(repl, src)

if out != src:
    html.write_text(out, encoding="utf-8")

for name, digest in sorted(seen.items()):
    print(f"  ✔ {name}?v={digest}")
for name in sorted(set(missing)):
    print(f"  ! {name} — δεν βρέθηκε στον φάκελο, το άφησα ως έχει", file=sys.stderr)

print("\nΤο index.html ενημερώθηκε." if out != src else "\nΌλα ήταν ήδη ενημερωμένα.")
PY
