import os
from pathlib import Path
from capstone import Cs, CS_ARCH_ARM, CS_MODE_THUMB

# === 1. 設定資料夾路徑與載入位址 ===
# 你說 bin 一律放在這個資料夾
FOLDER = r"C:\_python 2024\2024final"
LOAD_ADDR = 0x08000000  # STM32 Cortex-M 通常是 0x08000000，若不同再改

# === 2. 請使用者輸入檔名，不管叫什麼都可以 ===
print("bin 檔請放在資料夾：", FOLDER)
bin_name = input("請輸入要反組譯的 bin 檔檔名（例如：TGCreader_bootload.bin）：").strip()

BIN_PATH = Path(FOLDER) / bin_name

if not BIN_PATH.is_file():
    raise FileNotFoundError(f"找不到檔案：{BIN_PATH}")

# 自動產生輸出檔名：同檔名 + _disasm.txt
# 例如：abc.bin → abc_disasm.txt
OUTPUT_PATH = BIN_PATH.with_name(BIN_PATH.stem + "_disasm.txt")

# === 3. 建立 Capstone 反組譯器 (ARM Thumb 模式) ===
md = Cs(CS_ARCH_ARM, CS_MODE_THUMB)
md.detail = False  # 先簡單看指令就好，之後要看 operand 詳細再改 True

# === 4. 讀取 bin 檔 ===
with open(BIN_PATH, "rb") as f:
    data = f.read()

print("Disassembling:", BIN_PATH)
print("Total bytes:", len(data))
print("Output file:", OUTPUT_PATH)

# === 5. 反組譯並同時輸出到螢幕 & 檔案 ===
with open(OUTPUT_PATH, "w", encoding="utf-8") as out:
    out.write(f"; Disassembly of {BIN_PATH}\n")
    out.write(f"; Load address: 0x{LOAD_ADDR:08X}\n")
    out.write(f"; Total bytes : {len(data)}\n\n")

    for insn in md.disasm(data, LOAD_ADDR):
        line = "0x{addr:08X}:\t{mnemonic:<8} {op_str}".format(
            addr=insn.address,
            mnemonic=insn.mnemonic,
            op_str=insn.op_str
        )
        print(line)
        out.write(line + "\n")

print("Done. 反組譯結果已寫入：", OUTPUT_PATH)


