# -*- coding: utf-8 -*-
r"""Catalogue chi tiết 1 trang / 1 sản phẩm — Nệm Việt Nhật.
Hero (ảnh chính + thông số + giá) + bảng giá đầy đủ, mỗi nhóm ĐỘ DÀY có 1 ảnh bên trái.
Font SVN-Gilroy bundle sẵn trong skill (không phụ thuộc font đã cài).

API:
    from catalogue import build_catalogue
    build_catalogue(
        xlsx_path = r"...Bảng giá.xlsx",     # sheet đầu: HÌNH ẢNH|TÊN SP|ĐỘ DÀY|SKUs|GIÁ NIÊM YẾT|CHIẾT KHẤU|GIÁ BÁN LẺ
        product_key = "Standard",             # khớp 1 phần tên sản phẩm trong cột TÊN SP
        hero_image  = r"...\STANDARD.png",    # ảnh chính (hero)
        group_images = [                      # ảnh cho từng nhóm độ dày (path, caption); thiếu thì lặp lại
            (r"...\Cấu-tạo.png","CẤU TẠO BÊN TRONG"),
            (r"...\ChiTiet.jpg","CHI TIẾT CHẤT LIỆU"),
            (r"...\KhongGian.jpg","TRONG KHÔNG GIAN")],
        out_path = "Catalogue_Standard.png")
Yêu cầu: pip install python-pptx? KHÔNG — chỉ cần openpyxl + Pillow.
"""
import os, openpyxl
from PIL import Image, ImageDraw, ImageFont

_HERE=os.path.dirname(os.path.abspath(__file__))
FONTS=os.path.join(_HERE,"..","assets","fonts")
LOGO_WHITE=os.path.join(_HERE,"..","assets","logo_full_white.png")

GREEN=(0x40,0xA6,0x46); GREEN_DK=(0x2A,0x78,0x3B); DEEP=(0x05,0x34,0x21); LEMON=(0xFF,0xFD,0xA4)
PINK, PINK_DK, DARKPINK = GREEN, GREEN_DK, DEEP   # alias code cũ
BLACK=(0x23,0x1F,0x20); WHITE=(255,255,255); INK=(0x3A,0x35,0x37); MUTE=(0x8a,0x80,0x84)
CARD=(0xEA,0xF6,0xEC); CARD2=(0xF4,0xF5,0xF4); LINEC=(0xD9,0xE8,0xDC); PALE=(0xCD,0xEF,0xD3); GRAYTX=(0x9a,0x92,0x96)

def _font_dirs():
    import platform
    dirs = [FONTS]
    sysname = platform.system()
    if sysname == "Darwin":
        dirs += [os.path.expanduser("~/Library/Fonts"), "/Library/Fonts"]
    elif sysname == "Windows":
        dirs += [os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
                 os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")]
    else:
        dirs += [os.path.expanduser("~/.local/share/fonts/SVN-Gilroy"), os.path.expanduser("~/.local/share/fonts"), "/usr/share/fonts"]
    return [d for d in dirs if d and os.path.isdir(d)]

def _find_font(fname):
    """Tìm file font theo tên (không phân biệt hoa/thường, bỏ đuôi " (1)"): ưu tiên assets/fonts, rồi font đã cài trên máy."""
    want = os.path.splitext(fname)[0].lower().replace(" (1)", "")
    for d in _font_dirs():
        for root, _, files in os.walk(d):
            for f in files:
                if os.path.splitext(f)[0].lower().replace(" (1)", "") == want and f.lower().endswith((".otf", ".ttf")):
                    return os.path.join(root, f)
    raise FileNotFoundError(
        f"Không tìm thấy font '{fname}'. Bộ SVN-Gilroy có bản quyền, KHÔNG kèm trong bản công khai: "
        f"lấy bộ font nội bộ (Drive/Zalo phòng Marketing) rồi chép vào {FONTS} hoặc cài vào máy (scripts/install_fonts.py).")

def _f(fname,size): return ImageFont.truetype(_find_font(fname),size)
fb =lambda s:_f("SVN-GILROY BOLD.OTF",s)
fbl=lambda s:_f("SVN-GILROY BLACK.OTF",s)
fm =lambda s:_f("SVN-GILROY MEDIUM.OTF",s)
fsb=lambda s:_f("SVN-GILROY SEMIBOLD (1).TTF",s)
def money(v): return f"{int(round(v)):,}".replace(",",".")+"đ"

def _norm(v):
    import unicodedata
    s = unicodedata.normalize("NFD", str(v or "")); s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.upper().replace("Đ","D")

# Nhận cột theo TÊN TIÊU ĐỀ (không phụ thuộc vị trí). Mỗi cột: danh sách từ khoá, khớp 1 là đủ.
_COLS = {
    "name": ["SAN PHAM", "TEN SP", "TEN SAN PHAM", "PRODUCT"],
    "th":   ["DO DAY", "DAY"],
    "sz":   ["KICH THUOC", "SIZE", "KHO"],
    "ni":   ["NIEM YET", "GIA GOC", "LIST"],
    "ck":   ["CHIET KHAU", "CK", "GIAM"],
    "ba":   ["BAN LE", "GIA BAN", "KHUYEN MAI", "GIA KM", "SELL"],
}

def _find_header(ws, max_scan=15):
    """Dò dòng tiêu đề trong 15 dòng đầu: dòng nào có đủ tên SP + kích thước + giá bán lẻ. Trả (row, {key: col})."""
    for r in range(1, min(ws.max_row, max_scan) + 1):
        cells = {c: _norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)}
        found = {}
        for key, kws in _COLS.items():
            for c, txt in cells.items():
                if txt and any(k in txt for k in kws) and c not in found.values():
                    found[key] = c; break
        if {"name", "sz", "ba"} <= set(found):
            return r, found
    return None, None

def _read(xlsx_path, sheet=0):
    """Đọc bảng giá. Tự dò tiêu đề theo tên cột; không dò được thì dùng bố cục cũ
    (dòng 3+, B=TÊN SP, C=ĐỘ DÀY, D=KÍCH THƯỚC, E=NIÊM YẾT, F=CK, G=BÁN LẺ). Tên SP forward-fill."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True); ws = wb[wb.sheetnames[sheet]]
    hdr, col = _find_header(ws)
    if hdr is None:
        hdr, col = 2, {"name": 2, "th": 3, "sz": 4, "ni": 5, "ck": 6, "ba": 7}
    def val(r, key):
        c = col.get(key); return ws.cell(r, c).value if c else None
    prods = {}; cur = None; cur_th = ""
    for r in range(hdr + 1, ws.max_row + 1):
        nm = val(r, "name")
        if nm and str(nm).strip():
            cur = str(nm).replace("\n", " ").strip(); prods.setdefault(cur, []); cur_th = ""
        if cur is None: continue
        th = val(r, "th"); sz = val(r, "sz"); ni = val(r, "ni"); ck = val(r, "ck"); ba = val(r, "ba")
        if th and str(th).strip(): cur_th = str(th).strip()
        if sz and ba:
            try: ni = float(ni) if ni not in (None, "") else None; ck = float(ck) if ck not in (None, "") else None; ba = float(ba)
            except (TypeError, ValueError): continue
            prods[cur].append((th and str(th).strip() or "", str(sz).strip(), ni, ck, ba))
    return {k: v for k, v in prods.items() if v}
# Cam kết theo DÒNG (CHUNG/chinh-sach.md, anh Đạt chốt): FOAM ngủ thử 7 ngày · BH 10 năm (Enzo 12);
# CAO SU ngủ thử 30 ngày · BH 15 năm 1 đổi 1. Chọn bằng line="foam" | "caosu" | "enzo", hoặc truyền badges=.
BADGES_BY_LINE = {
    "foam":  ["Bảo hành 10 năm", "Ngủ thử 7 ngày", "Giao miễn phí toàn quốc"],
    "enzo":  ["Bảo hành 12 năm", "Ngủ thử 7 ngày", "Giao miễn phí toàn quốc"],
    "caosu": ["Bảo hành 15 năm 1 đổi 1", "Ngủ thử 30 ngày", "Giao miễn phí toàn quốc"],
}
DEFAULT_BADGES = None   # None = bắt buộc truyền line= hoặc badges= (không in mù)
DEFAULT_FOOTER = ["www.nemvietnhat.com", "Hotline 1800.646.810 (miễn cước)", "Nhận hàng ưng ý mới thanh toán"]

def build_catalogue(xlsx_path, product_key, hero_image, group_images=None, out_path=None, sheet=0,
                    badges=None, footer_items=None, year=None, subtitle="Dòng nệm chính hãng Nệm Việt Nhật",
                    line=None):
    """badges: 3 nhãn cam kết dưới thông số (mặc định DEFAULT_BADGES — số năm BẢO HÀNH phải xác nhận theo
    từng dòng trước khi in); footer_items: các mục dải chân trang (mặc định DEFAULT_FOOTER, hotline lấy từ
    brand guidelines); year: năm trên header (mặc định năm hiện tại)."""
    import datetime
    if badges is None:
        if line is None or line not in BADGES_BY_LINE:
            raise ValueError("Phải truyền line='foam'|'enzo'|'caosu' hoặc badges=[...]: cam kết bảo hành/ngủ thử khác theo dòng, không in mù.")
        badges = list(BADGES_BY_LINE[line])
    badges = list(badges)
    footer_items = list(footer_items) if footer_items else list(DEFAULT_FOOTER)
    year = year or datetime.date.today().year
    prods=_read(xlsx_path,sheet)
    pname=next((p for p in prods if product_key.lower() in p.lower()), None)
    if not pname: raise ValueError(f"Không tìm thấy sản phẩm khớp '{product_key}'. Có: {list(prods)}")
    rows=prods[pname]
    groups={}; order=[]; ct=None
    for th,sz,ni,ck,ba in rows:
        if th: ct=th
        if ct not in groups: groups[ct]=[]; order.append(ct)
        groups[ct].append((sz,ni,ck,ba))
    thicks=order; minban=min(x[4] for x in rows); sc=len(groups[thicks[0]])
    if out_path is None: out_path=f"Catalogue_{product_key.replace(' ','')}.png"

    W,H=1654,2339; page=Image.new("RGB",(W,H),WHITE); d=ImageDraw.Draw(page); M=60
    def wrap(text,font,maxw):
        out=[]; cur=""
        for w in text.split():
            t=(cur+" "+w).strip()
            if d.textlength(t,font=font)<=maxw: cur=t
            else: out.append(cur); cur=w
        if cur: out.append(cur)
        return out
    def fit(img_path,bx,by,bw,bh,radius=18,bg=WHITE):
        d.rounded_rectangle([bx,by,bx+bw,by+bh],radius=radius,fill=bg)
        if img_path and os.path.exists(img_path):
            im=Image.open(img_path).convert("RGBA"); pad=12
            s=min((bw-2*pad)/im.width,(bh-2*pad)/im.height); nw,nh=int(im.width*s),int(im.height*s)
            page.paste(im.resize((nw,nh),Image.LANCZOS),(bx+(bw-nw)//2,by+(bh-nh)//2),im.resize((nw,nh),Image.LANCZOS))
    def img_cap(img_path,bx,by,bw,bh,caption):
        d.rounded_rectangle([bx,by,bx+bw,by+bh],radius=14,fill=WHITE,outline=LINEC,width=2)
        if img_path and os.path.exists(img_path):
            im=Image.open(img_path).convert("RGBA"); pad=8
            s=min((bw-2*pad)/im.width,(bh-2*pad-30)/im.height); nw,nh=int(im.width*s),int(im.height*s)
            page.paste(im.resize((nw,nh),Image.LANCZOS),(bx+(bw-nw)//2,by+pad+((bh-30)-2*pad-nh)//2),im.resize((nw,nh),Image.LANCZOS))
        d.rounded_rectangle([bx,by+bh-34,bx+bw,by+bh],radius=14,fill=DARKPINK)
        d.rectangle([bx,by+bh-34,bx+bw,by+bh-14],fill=DARKPINK)
        d.text((bx+bw/2,by+bh-17),caption,font=fsb(19),fill=WHITE,anchor="mm")

    # header
    HB=190
    for x in range(W):
        t=x/W; u=t/0.5 if t<0.5 else (t-0.5)/0.5; a,b=(DARKPINK,PINK_DK) if t<0.5 else (PINK_DK,PINK)
        d.line([(x,0),(x,HB)],fill=tuple(int(a[k]+(b[k]-a[k])*u) for k in range(3)))
    if os.path.exists(LOGO_WHITE):
        lg=Image.open(LOGO_WHITE).convert("RGBA"); lh=78; lw=int(lg.width*lh/lg.height)
        page.paste(lg.resize((lw,lh),Image.LANCZOS),(M,(HB-lh)//2),lg.resize((lw,lh),Image.LANCZOS))
    d.text((W-M,58),f"CATALOGUE SẢN PHẨM {year}",font=fb(34),fill=WHITE,anchor="ra")
    d.text((W-M,105),"Bảng giá bán lẻ · Nệm Việt Nhật",font=fm(21),fill=PALE,anchor="ra")

    # hero
    hy=225; Hh=400; mw=520
    fit(hero_image,M,hy,mw,Hh,radius=22,bg=CARD)
    rx=M+mw+44; rw=W-M-rx; ty=hy+2
    for ln in wrap(pname,fb(48),rw): d.text((rx,ty),ln,font=fb(48),fill=PINK_DK,anchor="la"); ty+=59
    ty+=4; d.text((rx,ty),subtitle,font=fm(26),fill=MUTE,anchor="la"); ty+=54
    for k,v in [("Độ dày"," · ".join(thicks)),
                ("Kích thước",f"{groups[thicks[0]][0][0]} → {groups[thicks[0]][-1][0]} ({sc} size)"),
                ("Chiết khấu",f"đến {round(max(x[3] for x in rows)*100)}%")]:
        d.text((rx,ty),k,font=fm(26),fill=MUTE,anchor="la"); d.text((rx+210,ty),v,font=fsb(27),fill=BLACK,anchor="la"); ty+=48
    ty+=10; px=rx
    for p in badges:
        pw=d.textlength(p,font=fsb(23))+46
        if px+pw>W-M: px=rx; ty+=58
        d.rounded_rectangle([px,ty,px+pw,ty+48],radius=24,fill=WHITE,outline=PINK,width=2)
        d.text((px+pw/2,ty+24),p,font=fsb(23),fill=PINK_DK,anchor="mm"); px+=pw+14
    ty+=74
    d.text((rx,ty+12),"Chỉ từ",font=fm(27),fill=MUTE,anchor="la")
    d.text((rx+120,ty-12),money(minban),font=fbl(58),fill=PINK,anchor="la")

    # bảng giá
    taby=hy+Hh+34
    d.text((M,taby),"BẢNG GIÁ CHI TIẾT",font=fb(36),fill=BLACK,anchor="la")
    d.text((W-M,taby+12),"Đơn vị: VNĐ",font=fm(23),fill=MUTE,anchor="ra"); taby+=58
    imgW=348; rows_x0=M+imgW+22
    cx_size=rows_x0+8; cx_niem=rows_x0+610; cx_ck=rows_x0+810; cx_ban=W-M-22
    d.rounded_rectangle([M,taby,W-M,taby+54],radius=10,fill=BLACK)
    d.text((M+26,taby+27),"HÌNH ẢNH / KÍCH THƯỚC",font=fsb(25),fill=WHITE,anchor="lm")
    d.text((cx_niem,taby+27),"GIÁ NIÊM YẾT",font=fsb(25),fill=WHITE,anchor="rm")
    d.text((cx_ck,taby+27),"CK",font=fsb(25),fill=WHITE,anchor="mm")
    d.text((cx_ban,taby+27),"GIÁ BÁN LẺ",font=fsb(25),fill=WHITE,anchor="rm")
    y=taby+54
    gi_imgs = group_images or []
    gh=54; ggap=12; n=len(thicks); nrows=sum(len(groups[t]) for t in thicks)
    rh=int((H-118-30 - y - n*gh - (n-1)*ggap)/nrows)   # giãn đều cho đầy trang
    grpcol=[PINK,PINK_DK,DARKPINK]
    for gi,t in enumerate(thicks):
        d.rectangle([M,y,W-M,y+gh],fill=grpcol[gi%3])
        d.text((M+26,y+gh//2),f"ĐỘ DÀY {t}",font=fb(28),fill=WHITE,anchor="lm")
        d.text((cx_ban,y+gh//2),f"{len(groups[t])} kích thước",font=fm(22),fill=PALE,anchor="rm")
        y+=gh; blk_h=len(groups[t])*rh
        if gi_imgs:
            ip,cap=gi_imgs[gi%len(gi_imgs)]; img_cap(ip,M+6,y+8,imgW-12,blk_h-16,cap)
        for ri,(sz,ni,ck,ba) in enumerate(groups[t]):
            ry=y+ri*rh
            if ri%2==1: d.rectangle([rows_x0,ry,W-M,ry+rh],fill=CARD2)
            d.text((cx_size,ry+rh//2),sz,font=fsb(26),fill=INK,anchor="lm")
            nt=money(ni); d.text((cx_niem,ry+rh//2),nt,font=fm(24),fill=GRAYTX,anchor="rm")
            ntw=d.textlength(nt,font=fm(24)); d.line([cx_niem-ntw,ry+rh//2,cx_niem,ry+rh//2],fill=GRAYTX,width=2)
            d.text((cx_ck,ry+rh//2),f"-{round(ck*100)}%",font=fsb(25),fill=PINK,anchor="mm")
            d.text((cx_ban,ry+rh//2),money(ba),font=fb(28),fill=PINK_DK,anchor="rm")
        y+=blk_h+ggap

    # footer
    fyy=H-104
    d.rounded_rectangle([M,fyy,W-M,fyy+62],radius=18,fill=DARKPINK)
    d.text((W//2,fyy+31),"   ·   ".join(footer_items),font=fsb(20),fill=WHITE,anchor="mm")
    d.text((W//2,fyy+86),"NÂNG TẦM CUỘC SỐNG",font=fsb(18),fill=PINK,anchor="mm")
    page.save(out_path,"PNG"); return out_path

if __name__=="__main__":
    print("Đây là module — import build_catalogue() để dùng. Xem docstring đầu file.")
