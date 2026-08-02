import os

folder = "csv file"  # 相对路径

if os.path.exists(folder):
    print(f"📂 文件夹 '{folder}' 中的文件列表：")
    files = os.listdir(folder)
    if files:
        for f in files:
            print(f"  - {f}")
    else:
        print("  （文件夹为空）")
else:
    print(f"❌ 文件夹不存在: {folder}")
    print("请确认当前运行目录是否为项目根目录。")

#folder = r"C:\Users\Shree Subi C\Documents\mini-project-s7\csv file"

#print("Files in folder:")
#print(os.listdir(folder))