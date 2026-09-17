---
name: nem-viet-nhat-brand
description: >-
  Bộ nhận diện thương hiệu Nệm Việt Nhật (màu xanh #40A646, font SVN Gilroy, logo ngôi sao, nền gradient
  xanh, tagline "Nâng Tầm Cuộc Sống") theo Brand Book 2026, kèm công cụ dựng slide PowerPoint và catalogue
  sản phẩm đúng chuẩn. BẮT BUỘC dùng skill này mỗi khi người dùng yêu cầu làm slide/deck/bài thuyết trình,
  file PowerPoint/PPTX, catalogue, hoặc thiết kế bất kỳ ấn phẩm nào (bìa, poster, tờ rơi, banner, danh thiếp,
  social) CHO Nệm Việt Nhật — kể cả khi không nói rõ "theo nhận diện". Cũng dùng khi cần tra mã màu, font,
  logo, tagline, vùng an toàn logo của Nệm Việt Nhật, hay khi cập nhật slide vào deck đã làm theo nhận diện
  này. KHÔNG dùng cho Nệm Thuần Việt (skill nem-thuan-viet-brand, màu hồng), không dùng cho kịch bản video TikTok.
---

# Nệm Việt Nhật — Nhận diện thương hiệu & dựng slide / catalogue

Skill giữ toàn bộ nhận diện **Nệm Việt Nhật** (Brand Book 2026) và bộ công cụ Python dựng slide PowerPoint,
catalogue đúng chuẩn. Cùng khung với skill `nem-thuan-viet-brand` (hồng) nhưng **màu, logo, nền, tagline, chính
sách khác hẳn** — đừng lấy nhầm asset hay số của Thuần Việt.

## Nhận diện cốt lõi (tóm tắt nhanh)

**Màu chính (Brand Book 2026):** Xanh `#40A646` (chủ đạo) · Xanh rêu `#2A783B` · Xanh đậm nền tối `#053421` ·
Vàng chanh `#FFFDA4` (chỉ cho tiêu đề lớn trên nền tối / đầu gradient phụ) · Đen `#231F20` · Trắng · Kem card `#FCFCF5`.
Gradient bìa/section: `#40A646 → #053421`. Chữ trên nền tối: tiêu đề **vàng chanh**, chữ phụ trắng.
File logo PNG gốc tô `#28A54A`/`#277714` (bảng màu 2025), lệch nhẹ với brand book — **anh Đạt đã chốt 17/09/2026: màu theo
brand book, logo giữ file gốc**, không tô lại, không hỏi lại. Chi tiết mục 3 của [references/brand-guidelines.md](references/brand-guidelines.md).

**Font:** SVN Gilroy — **SemiBold** (H1 + H2) / **Regular** (body); tiêu đề lớn in hoa trên nền tối dùng Bold/Heavy.
Trong `brand_pptx`, chọn nấc bằng `weight=` ở `txt()`/`bullets()`: `thin light regular medium semibold bold heavy black`
(+ `italic=True`), đủ 16 biến thể. `head()` mặc định kicker + tiêu đề SemiBold theo brand book. Mọi run tự khai font ở cả
3 khe latin/ea/cs (bắt buộc, kẻo dấu tiếng Việt rớt nấc); viết run tay thì gọi `set_font_slots(run, tên_font)`.
Bộ font 16 file bundle trong `assets/fonts/`; máy chưa cài: `python3 scripts/install_fonts.py` (macOS/Windows/Linux, không cần admin).

**Logo:** ngôi sao trắng trong vòng tròn xanh + wordmark NỆM VIỆT NHẬT + tagline. 4 phiên bản: ngang (chính) · dọc · biểu tượng · chữ.
Nền trắng/kem → logo xanh; nền xanh/gradient/đen → logo **trắng**; đơn sắc → đen. Vùng an toàn **4x**, cao tối thiểu 35px, đặt ở 4 góc hoặc
chính giữa cạnh. **6 lỗi cấm:** kéo méo · nền rối · đổ bóng · đổi khoảng cách chữ · thêm viền · đổi màu.

**Tagline:** *Nâng Tầm Cuộc Sống.* (Không có slogan khác; không lấy "101 đêm" của Thuần Việt.)

**Liên hệ in ấn phẩm:** www.nemvietnhat.com · Hotline **1800.646.810** (miễn cước) · H76–H77 Dương Thị Giang, P. Tân Thới Nhất, Q.12, TP.HCM (anh Đạt chốt 17/09/2026, cùng địa chỉ Thuần Việt).

## Luật nghiệp vụ khi ấn phẩm có số (sai = sai cam kết bán hàng)
- **Cam kết theo DÒNG, không dùng chung:** FOAM (Sora, Nano, Latex Gold — *Latex Gold là foam*) ngủ thử **7 ngày**, BH **10 năm**; Enzo BH **12 năm**;
  CAO SU (Natural, Ultimate, Lavender) ngủ thử **30 ngày**, BH **15 năm 1 đổi 1**. Nguồn: `~/Documents/Claude/CHUNG/chinh-sach.md`.
- Chung: giao miễn phí toàn quốc · "nhận hàng ưng ý mới thanh toán" (không viết "không phí COD") · tặng 7 món (ngủ thử chỉ cho nệm) · ISO 9001:2015 + Hợp Quy.
- Giá/size/chiết khấu tra `CHUNG/san-pham-gia.md`. Cấm chê đại lý, hứa chữa bệnh, "Nhật Bản chứng nhận", hứa giờ giao.

## Asset có sẵn (`assets/`)

| File | Dùng cho |
|---|---|
| `logo_full_white.png` | Logo ngang trắng — nền gradient / xanh / đen |
| `logo_full_transparent.png` | Logo ngang xanh, nền trong suốt — nền trắng/kem |
| `logo_full_horizontal.png` | Logo ngang xanh trên nền trắng |
| `logo_full_black.png` | Logo ngang đen — in đơn sắc |
| `logo_vertical_transparent.png` | Logo dọc xanh (bản phụ) |
| `logo_icon_white.png` / `logo_icon_transparent.png` / `logo_icon_mark.png` | Biểu tượng ngôi sao (trắng / xanh trong suốt / xanh nền trắng) — favicon, watermark |
| `bg_title.png` | Nền gradient bìa: xanh đậm, vệt sáng góc trên **phải**, watermark sao mờ bên phải (chữ đặt bên trái trên vùng tối) |
| `bg_section.png` | Nền gradient trang ngăn: vệt sáng góc trên **trái**, watermark sao bên trái |
| `bg_gradient.png` | Nền gradient trơn |

> **Đường dẫn thư mục skill:** các ví dụ dưới dùng `~/.claude/skills/nem-viet-nhat-brand/`. Nếu cài qua **chợ plugin**
> (`/plugin install ntv-nhan-dien@nem-thuan-viet`) thì skill nằm trong `~/.claude/plugins/…/ntv-nhan-dien/skills/nem-viet-nhat-brand/`.
> Cách chắc nhất: lấy thư mục chứa chính file SKILL.md này rồi trỏ tới `scripts/` bên cạnh, hoặc tìm bằng
> `find ~/.claude -path "*nem-viet-nhat-brand/scripts/brand_pptx.py" 2>/dev/null | head -1`.

## Dựng slide PowerPoint

Dùng `scripts/brand_pptx.py` — palette, asset, hàm dựng sẵn. Khung **16:9 (13.333 × 7.5 inch)**, toạ độ inch.
**Chuẩn bị** (một lần): `python3 -m pip install python-pptx Pillow`

```python
import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/nem-viet-nhat-brand/scripts"))
from brand_pptx import *

prs = new_deck(); set_doc_title("Tên Tài Liệu · Nệm Việt Nhật")
s = add_slide(prs); cover(s, "TÀI LIỆU NỘI BỘ", "TIÊU ĐỀ\nLỚN", "Phụ đề", "Phòng ban · Nệm Việt Nhật")   # tiêu đề vàng chanh
s = add_slide(prs); divider(s, "PHẦN 01", "Tên Section", "Mô tả ngắn", "01")
s = add_slide(prs); solid_bg(s, WHITE); head(s, "01 · MỤC", "Tiêu Đề Nội Dung")
bullets(s, [("Ý chính: ", "diễn giải"), "ý thường"], 0.8, 1.9, 11.7, 4, 13, INK)
footer(s, "Tên section", 3)
prs.save("output.pptx")
```

**Quy ước bố cục:**
- Bìa & section & kết = nền gradient (`cover`/`divider`), logo trắng, **tiêu đề vàng chanh `LEMON`**, chữ phụ trắng.
- Trang nội dung = nền trắng, `head()` (kicker xanh + tiêu đề SemiBold đen), `footer()` dưới.
- Card: `rrect()` nền `CREAM`/`CARD`/`CARD2`; header tô `GREEN`/`GREEN_DK`/`DEEP`, chữ trắng; chữ trong card 11–13pt.
- Số liệu nổi bật: số lớn 36–72pt `GREEN` (weight `black`/`heavy`), nhãn nhỏ `GREEN_DK`.
- Không gạch chân tiêu đề dài, không thanh màu full-width; vàng chanh **chỉ** trên nền tối.
- Hằng màu: `GREEN GREEN_DK DEEP LEMON BLACK WHITE CREAM GRAY INK MUTE CARD CARD2 LINEC PALEGREEN ACCENTGR`
  (+ tên chung `PRIMARY PRIMARY_DK DARK ACCENT`). Hàm: `new_deck add_slide add_bg solid_bg cover divider head footer
  logo_corner logo_big txt bullets rrect oval chevron line set_font_slots`.

## Catalogue sản phẩm (ảnh + bảng giá)

`scripts/catalogue.py` → trang A4 dọc / 1 sản phẩm từ Excel. **Tự dò dòng tiêu đề và cột theo TÊN** (SẢN PHẨM/TÊN SP · ĐỘ DÀY ·
KÍCH THƯỚC · GIÁ NIÊM YẾT · CHIẾT KHẤU (0.25 = 25%) · GIÁ BÁN LẺ), vị trí cột nào cũng được, tên SP forward-fill theo dòng trống; không dò
được thì rơi về bố cục cũ (dòng 3+, cột B..G). Đọc thẳng file thật `MARKETING/Bảng giá Nệm Việt Nhật.xlsx` (đã test 17/09/2026, khớp
21/21 dòng Nano). `python3 -m pip install openpyxl Pillow`.

```python
from catalogue import build_catalogue
build_catalogue("Bảng giá.xlsx", product_key="Nano", hero_image="/…/nano.png",
    group_images=[("/…/cau-tao.png","CẤU TẠO BÊN TRONG"), ("/…/chi-tiet.jpg","CHI TIẾT CHẤT LIỆU")],
    out_path="Catalogue_Nano.png",
    line="foam")          # BẮT BUỘC: "foam" | "enzo" | "caosu" → nhãn bảo hành/ngủ thử đúng dòng; hoặc badges=[...] tự ghi
```
**Không có mặc định cam kết** — thiếu `line=`/`badges=` script báo lỗi, đúng ý: không in mù bảo hành. `footer_items` mặc định web + hotline
1800.646.810 + "Nhận hàng ưng ý mới thanh toán"; `year=` mặc định năm hiện tại; `subtitle=` tuỳ chỉnh. Gộp PNG thành PDF:
`img.convert("RGB").save("catalogue.pdf", save_all=True, append_images=[...])`.

## Kiểm tra (QA) — render slide ra ảnh để soi

```bash
python3 ~/.claude/skills/nem-viet-nhat-brand/scripts/qa_render.py output.pptx      # → <thư mục>/_qa/slide-01.png …
```
- **macOS:** có LibreOffice + poppler thì dùng; không có thì tự chuyển sang **Quick Look** (không cần cài gì; ép: `--quicklook`).
- **Linux:** LibreOffice + poppler. **Windows:** PowerPoint qua COM.
Đọc từng PNG, soi: tràn/đè/lệch · **chữ CÓ DẤU ở nấc đậm/mảnh** (ủ ố ắ ặ ỡ ữ ệ) · logo đúng bản trắng/xanh theo nền · vàng chanh không lọt
lên nền trắng. Sửa lỗi rõ ràng rồi dừng. Thiếu công cụ render → giao file kèm nói rõ **chưa QA hình**.

## Quy trình gợi ý
1. Nắm chủ đề, định dạng (mặc định .pptx 16:9), nguồn nội dung; **số liệu/cam kết** tra sổ chung `CHUNG/`, thiếu thì hỏi anh Đạt.
2. Viết `build_*.py` import `brand_pptx`, dựng từng slide. 3. `qa_render.py` → soi → sửa. 4. Giao file + nêu lưu ý font / điểm chưa xác nhận.
