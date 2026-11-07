"""Extract the .xlsx attachment from a Gmail get_message RAW JSON dump.
Usage: python extract_xlsx.py <raw.json> <out_dir>"""
import sys, json, base64, email, email.policy, pathlib

sys.stdout.reconfigure(encoding="utf-8")
src, out = sys.argv[1], pathlib.Path(sys.argv[2])
raw = json.loads(pathlib.Path(src).read_text(encoding="utf-8"))["raw"]
raw += "=" * (-len(raw) % 4)
msg = email.message_from_bytes(base64.urlsafe_b64decode(raw), policy=email.policy.default)
for part in msg.iter_attachments():
    name = part.get_filename() or ""
    if name.lower().endswith(".xlsx"):
        p = out / name
        p.write_bytes(part.get_content())
        print(p)
