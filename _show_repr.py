
path = r"C:\Users\Alberto\Desktop\Programa de control de flotilla automotriz para Win-Mac 210826\modulo_ventas.py"
with open(path, "r", encoding="utf-8") as f:
    lines = f.read().split("\n")
# print repr of lines 507..574 (0-indexed 506..573)
for i in range(506, 574):
    print(i+1, repr(lines[i]))
