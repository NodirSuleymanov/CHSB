# CHSB Ball Hisoblash Dasturi

Excel jurnal fayllaridan **CHSB** (Chorak bo'yicha Sinalish Bali) ustunidagi ballarni o'qib, har bir sinf bo'yicha:
- Har bir o'quvchining CHSB balini
- Jami (yig'indi) ballni
- O'rtacha ballni
- Eng yuqori va eng past ballni

avtomatik hisoblab chiqaradi.

---

## O'rnatish

```bash
pip install -r requirements.txt
```

---

## Ishlatish

### 1. Bitta fayl
```bash
python3 chsb_calculator.py jurnal.xlsx
```

### 2. Bir nechta fayl
```bash
python3 chsb_calculator.py sinf_7a.xlsx sinf_7b.xlsx sinf_8a.xlsx
```

### 3. Papkadagi barcha fayllar (avtomatik)
```bash
# Excel fayllarni dastur bilan bir papkaga qo'ying, keyin:
python3 chsb_calculator.py
```

---

## Natija namunasi

```
============================================================
Fayl: jurnal.xlsx
============================================================

  Varaq: 'Sinf jurnali'
  CHSB ustuni: 22-ustun, sarlavha qatori: 1

  №    O'quvchi ismi                   CHSB bali
  ---- ------------------------------ ----------
  1    Adhamov Elbek                        28.0
  2    Ahmedov Abdulatif                    20.0
  ...
  35   Shaxobiddinova Dilyoraxon            34.0

  ──────────────────────────────────────────────
  Jami o'quvchilar soni    : 35
  Ball kiritilgan o'quvchilar: 32
  Yig'indi (jami ball)      : 886.0
  O'rtacha ball             : 27.69
  Eng yuqori ball           : 35.0
  Eng past ball             : 17.0
```

---

## Eslatmalar

- CHSB ustuni nomi Excel faylning birinchi 10 qatorida bo'lishi kerak
- `s`, `k`, `dq` va boshqa harfli belgilar **hisobga olinmaydi** (faqat raqamlar)
- Bir Excel faylida bir nechta varaq bo'lsa, **barchasi** hisoblanadi
- `.xlsx`, `.xls`, `.xlsm` formatlar qo'llab-quvvatlanadi
