"""Excel(COM)로 xlsx 의 시트들을 PNG 로 내보내 눈으로 확인 — export_sheets.py <xlsx> <outdir> <시트명> ..."""
import sys, os, site
site.addsitedir(site.getusersitepackages())
sys.stdout.reconfigure(encoding="utf-8")
import pymupdf
import win32com.client

path, outdir, names = os.path.abspath(sys.argv[1]), sys.argv[2], sys.argv[3:]
os.makedirs(outdir, exist_ok=True)
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(path, 0, True)
    xl.CalculateFull()
    for name in names:
        ws = wb.Worksheets(name)
        ws.Visible = -1
        ws.Activate()
        ps = ws.PageSetup
        ps.Orientation = 2
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = False
        pdf = os.path.join(outdir, f"{name}.pdf")
        ws.ExportAsFixedFormat(0, pdf)
        doc = pymupdf.open(pdf)
        for i, pg in enumerate(doc):
            if i >= 2:
                break
            fn = os.path.join(outdir, f"{name}_{i + 1}.png")
            pg.get_pixmap(matrix=pymupdf.Matrix(1.6, 1.6)).save(fn)
            print(fn)
        doc.close()
    wb.Close(False)
finally:
    xl.Quit()
