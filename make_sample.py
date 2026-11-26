# -*- coding: utf-8 -*-
"""검수 도구 테스트용 샘플 상품 엑셀 생성 (정상 행 + 의도적 오류 행, 셀 위 이미지 포함)"""
import io, random, sys
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment
from PIL import Image, ImageDraw

random.seed(7)
out = sys.argv[1] if len(sys.argv) > 1 else "sample_products.xlsx"

HEAD = ["상품코드", "상품명", "브랜드", "카테고리", "상품이미지", "이미지URL", "옵션",
        "단가", "상시가", "할인가", "할인율", "재고", "판매시작일", "판매종료일",
        "배송비", "원산지", "상세설명", "비고"]

items = [
    ("유기농 현미 4kg", "바른들", "식품>쌀/잡곡"), ("무농약 사과 3kg", "청송농원", "식품>과일"),
    ("한우 등심 500g", "한마음축산", "식품>정육"), ("제주 감귤 5kg", "탐라팜", "식품>과일"),
    ("국산 참기름 350ml", "고소방앗간", "식품>양념"), ("저당 그래놀라 500g", "모닝그린", "식품>시리얼"),
    ("무선 청소기 V12", "클린텍", "가전>청소기"), ("스테인리스 텀블러 600ml", "데일리컵", "주방>텀블러"),
    ("면 100% 수건 5매", "코튼데이", "생활>욕실"), ("대나무 칫솔 10입", "에코라이프", "생활>욕실"),
    ("LED 스탠드 조명", "브라이트", "가전>조명"), ("캠핑 접이식 의자", "아웃핏", "레저>캠핑"),
    ("여성 경량 패딩", "윈터웨어", "패션>아우터"), ("남성 기모 슬랙스", "모던핏", "패션>하의"),
    ("어린이 우산", "키즈레인", "잡화>우산"), ("요가 매트 8mm", "핏앤밸런스", "레저>요가"),
    ("원두커피 1kg 블렌드", "로스터리K", "식품>커피"), ("유아 물티슈 100매x10", "베이비순", "유아>물티슈"),
    ("전기 그릴 팬", "쿡마스터", "주방>조리가전"), ("멀티탭 6구 3m", "세이프라인", "가전>전기용품"),
    ("수제 딸기잼 300g", "잼공방", "식품>잼"), ("고양이 사료 2kg", "펫프렌즈", "반려>사료"),
    ("블루투스 이어폰", "사운드온", "가전>음향"), ("실리콘 주걱 3종", "키친픽", "주방>도구"),
    ("홍삼 스틱 30포", "정관원", "건강>홍삼"), ("비타민C 1000 120정", "헬시랩", "건강>비타민"),
    ("양털 슬리퍼", "포근", "잡화>슬리퍼"), ("방수 백팩 25L", "트레일", "잡화>가방"),
    ("규조토 발매트", "드라이홈", "생활>욕실"), ("천연 수세미 3입", "에코라이프", "주방>세척"),
]
COLORS = ["#E57373", "#64B5F6", "#81C784", "#FFB74D", "#BA68C8", "#4DB6AC", "#A1887F", "#90A4AE"]

def make_png(label, color, w=160, h=160):
    im = Image.new("RGB", (w, h), color)
    d = ImageDraw.Draw(im)
    d.rectangle([6, 6, w - 7, h - 7], outline="white", width=3)
    d.text((14, h // 2 - 8), label, fill="white")
    b = io.BytesIO(); im.save(b, "PNG"); b.seek(0)
    return b

wb = Workbook(); ws = wb.active; ws.title = "상품목록"
ws.append(HEAD)
for c in ws[1]:
    c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
    c.alignment = Alignment(horizontal="center")

# 행 번호(데이터 기준 1부터) -> 심어 둘 오류
ERRORS = {
    3: "할인가>상시가", 6: "역마진", 9: "할인율 불일치", 12: "가격 문자", 15: "가격 누락",
    18: "이미지 없음", 21: "상품명 누락", 24: "할인율 과다", 27: "할인 없음", 29: "음수 가격",
}

start = date(2026, 12, 1)
for i, (name, brand, cat) in enumerate(items, start=1):
    cost = random.randrange(30, 400) * 100
    regular = int(round(cost * random.uniform(1.5, 2.2), -2))
    rate = random.choice([10, 15, 20, 25, 30])
    sale = int(round(regular * (100 - rate) / 100, -2))
    stated = rate / 100
    err = ERRORS.get(i)
    if err == "할인가>상시가": sale = regular + 3000
    if err == "역마진": sale = cost - 1000
    if err == "할인율 불일치": stated = 0.40
    if err == "가격 문자": regular = "문의"
    if err == "가격 누락": sale = None
    if err == "상품명 누락": name = ""
    if err == "할인율 과다": sale = int(round(regular * 0.08, -2)); stated = round(1 - sale / regular, 2)
    if err == "할인 없음": sale = regular; stated = 0
    if err == "음수 가격": cost = -cost
    s = start + timedelta(days=random.randrange(0, 20))
    row = [f"P{1000 + i}", name, brand, cat, None,
           f"https://picsum.photos/seed/p{1000 + i}/400/400.jpg" if i % 3 == 0 and err != "이미지 없음" else "",
           random.choice(["단일상품", "색상: 블랙/화이트", "사이즈: S/M/L", "용량 선택", "-"]),
           cost, regular, sale, stated, random.randrange(0, 500), s, s + timedelta(days=random.randrange(7, 60)),
           random.choice([0, 3000, 2500]), random.choice(["국산", "중국", "베트남", "국산(제주)"]),
           f"{brand}의 {name or '상품'}입니다. 선물용으로도 좋습니다.", err and f"[테스트오류] {err}" or ""]
    ws.append(row)
    r = i + 1
    ws.row_dimensions[r].height = 62
    for col in ("H", "I", "J", "O"): ws[f"{col}{r}"].number_format = "#,##0"
    ws[f"K{r}"].number_format = "0%"
    for col in ("M", "N"): ws[f"{col}{r}"].number_format = "yyyy-mm-dd"
    if err != "이미지 없음":
        img = XLImage(make_png(f"P{1000 + i}", COLORS[i % len(COLORS)]))
        img.width, img.height = 80, 80
        ws.add_image(img, f"E{r}")
        if i % 5 == 0:  # 일부 상품은 이미지 2장
            img2 = XLImage(make_png(f"P{1000 + i}-2", COLORS[(i + 3) % len(COLORS)]))
            img2.width, img2.height = 80, 80
            ws.add_image(img2, f"G{r}")

widths = [10, 24, 12, 14, 13, 30, 16, 10, 10, 10, 8, 7, 12, 12, 8, 10, 40, 22]
for idx, w in enumerate(widths):
    ws.column_dimensions[chr(65 + idx)].width = w

ws2 = wb.create_sheet("메모"); ws2.append(["이 시트는 시트 선택 테스트용입니다."])
wb.save(out)
print("saved", out)
