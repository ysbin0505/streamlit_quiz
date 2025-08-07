import os
dir_a = "/Users/data.ly/Desktop/말뭉치배포/신문/A"
dir_b = "/Users/data.ly/Desktop/말뭉치배포/신문/B"
print("A exists:", os.path.isdir(dir_a))
print("B exists:", os.path.isdir(dir_b))
print("A list:", os.listdir(dir_a) if os.path.isdir(dir_a) else "없음")
print("B list:", os.listdir(dir_b) if os.path.isdir(dir_b) else "없음")
