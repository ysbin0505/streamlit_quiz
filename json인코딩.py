import json

with open("/Users/data.ly/Downloads/project_176.json", "rb") as f:
    orig = f.read()
print("원본:", len(orig))

obj = json.loads(orig.decode("utf-8"))

utf8_pretty = json.dumps(obj, ensure_ascii=False, indent=2).encode("utf-8")
ascii_pretty = json.dumps(obj, ensure_ascii=True, indent=2).encode("utf-8")
compact      = json.dumps(obj, ensure_ascii=False, separators=(",", ":")).encode("utf-8")

print("UTF-8 pretty:", len(utf8_pretty))
print("\\uXXXX pretty:", len(ascii_pretty))
print("compact:", len(compact))
