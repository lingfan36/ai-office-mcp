import win32com.client, time, tempfile, os, openpyxl

path = os.path.join(tempfile.gettempdir(), f"val_diag_{int(time.time())}.xlsx")
wb0 = openpyxl.Workbook()
ws0 = wb0.active
ws0["A1"] = "test"
wb0.save(path)

app = win32com.client.Dispatch("Excel.Application")
app.Visible = True
wb = app.Workbooks.Open(path)
ws = wb.Worksheets(1)

tests = [
    ("ASCII list",    "B1:B5", 3, "a,b,c"),
    ("Unicode list",  "B1:B5", 3, "资产,负傺,权益"),  # 资产,负债,权益
    ("Range ref",     "B1:B5", 3, "=$D$1:$D$3"),
    ("Whole number",  "C1:C5", 1, "0"),   # xlValidateWholeNumber=1
    ("Decimal >=0",   "C1:C5", 4, "0"),   # xlValidateDecimal=4
]

ws.Range("D1:D3").Value = [["资产"], ["负债"], ["权益"]]

for label, addr, vtype, f1 in tests:
    rng = ws.Range(addr)
    rng.Validation.Delete()
    last_err = None
    for call_kwargs in [
        {"Type": vtype, "Formula1": f1},
        {"Type": vtype, "Formula1": f1, "Operator": 1},
        {"Type": vtype, "AlertStyle": 1, "Formula1": f1},
    ]:
        try:
            rng.Validation.Add(**call_kwargs)
            print(f"OK  {label}  kwargs={list(call_kwargs.keys())}")
            last_err = None
            break
        except Exception as e:
            last_err = e
    if last_err is not None:
        print(f"NG  {label}  all attempts failed  err={last_err}")
    rng.Validation.Delete()

wb.Close(SaveChanges=False)
print("diag done")
